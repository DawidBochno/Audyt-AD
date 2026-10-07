#!/usr/bin/env python3
"""Audyt AD - konta uzytkownikow i komputery, ktore nie logowaly sie do
domeny od N dni. Tylko odczyt (ADSI/LDAP), nic nie zmienia w domenie.
Wynik w oknie i w pliku XLSX (zakladki Konta i Komputery).

Uruchomienie: python audyt_ad.py            (GUI)
              python audyt_ad.py --selftest (test logiki)
"""
import datetime
import os
import re
import sys
import threading

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

FILTRY = {
    "Konta": "(&(objectCategory=person)(objectClass=user))",
    "Komputery": "(objectCategory=computer)",
}
ATRYBUTY = ["sAMAccountName", "name", "displayName", "operatingSystem",
            "lastLogonTimestamp", "whenCreated", "userAccountControl",
            "distinguishedName"]
KOLUMNY = {
    "Konta": ["Login", "Imie i nazwisko", "Ostatnie logowanie", "Dni bez logowania",
              "Utworzone", "Stan", "Jednostka (OU)"],
    "Komputery": ["Nazwa", "System", "Ostatnie logowanie", "Dni bez logowania",
                  "Utworzony", "Stan", "Jednostka (OU)"],
}
UAC_WYLACZONE = 0x2
FILETIME_1970 = 116444736000000000

# ------------------------------------------------------------- konwersje ----


def duza_liczba(v):
    """IADsLargeInteger (HighPart/LowPart) z ADO albo int -> int / None."""
    if v is None:
        return None
    if isinstance(v, int):
        return v
    return (v.HighPart << 32) + (v.LowPart & 0xFFFFFFFF)


def filetime(n):
    """FILETIME (100 ns od 1601, UTC) -> lokalny datetime; 0 / 'nigdy' -> None."""
    if not n or n >= 0x7FFFFFFFFFFFFFFF:
        return None
    return datetime.datetime.fromtimestamp((n - FILETIME_1970) / 1e7)


def lokalny(v):
    """Data z ADO (UTC, czasem bez strefy) -> lokalny datetime bez strefy."""
    if v is None:
        return None
    if v.tzinfo is None:
        v = v.replace(tzinfo=datetime.timezone.utc)
    return datetime.datetime.fromtimestamp(v.timestamp())


def jednostka(dn):
    """'CN=Jan,OU=Kadry,OU=Urzad,DC=urzad,DC=local' -> 'Urzad/Kadry'."""
    czesci = re.split(r"(?<!\\),", dn or "")[1:]
    return "/".join(c.split("=", 1)[1] for c in reversed(czesci)
                    if c.upper().startswith(("OU=", "CN=")))


def baza(domena):
    return ",".join("DC=" + d for d in domena.split("."))


def obiekt(rodzaj, r):
    """Wiersz z AD (slownik atrybutow) -> rekord do raportu."""
    pierwszy = lambda v: v[0] if isinstance(v, (tuple, list)) and v else v
    r = {k: pierwszy(v) for k, v in r.items()}
    return {
        "nazwa": (r.get("sAMAccountName") if rodzaj == "Konta" else r.get("name")) or "",
        "opis": (r.get("displayName") if rodzaj == "Konta" else r.get("operatingSystem")) or "",
        "ostatnie": filetime(duza_liczba(r.get("lastLogonTimestamp"))),
        "utworzone": lokalny(r.get("whenCreated")),
        "wylaczone": bool((r.get("userAccountControl") or 0) & UAC_WYLACZONE),
        "ou": jednostka(r.get("distinguishedName")),
    }

# ---------------------------------------------------------------- AD ----


def pobierz(domena, login="", haslo=""):
    """Odczyt kont i komputerow z AD przez ADSI (ADO). Bez loginu - jako
    zalogowany uzytkownik Windows. Zwraca {"Konta": [...], "Komputery": [...]}."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()  # wywolanie z watku GUI
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Provider = "ADsDSOObject"
    if login:
        conn.Properties("User ID").Value = login
        conn.Properties("Password").Value = haslo
        conn.Properties("Encrypt Password").Value = True
    conn.Open("Active Directory Provider")
    wynik = {}
    for rodzaj, filtr in FILTRY.items():
        cmd = win32com.client.Dispatch("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.Properties("Page Size").Value = 1000  # bez tego max 1000 wynikow
        cmd.CommandText = "<LDAP://%s/%s>;%s;%s;subtree" % (
            domena, baza(domena), filtr, ",".join(ATRYBUTY))
        rs = cmd.Execute()
        rs = rs[0] if isinstance(rs, tuple) else rs  # (Recordset, RecordsAffected)
        wiersze = []
        while not rs.EOF:
            wiersze.append(obiekt(rodzaj, {a: rs.Fields(a).Value for a in ATRYBUTY}))
            rs.MoveNext()
        wynik[rodzaj] = wiersze
    conn.Close()
    return wynik

# ------------------------------------------------------------- analiza ----


def nieaktywne(obiekty, dni, teraz, wylaczone=False):
    """Obiekty bez logowania od `dni` dni. Nigdy nie logowane - tylko jesli
    utworzone wczesniej niz `dni` dni temu (nowe konto to nie martwe konto)."""
    granica = teraz - datetime.timedelta(days=dni)
    out = []
    for o in obiekty:
        if o["wylaczone"] and not wylaczone:
            continue
        ost = o["ostatnie"]
        if (ost or o["utworzone"] or teraz) > granica:
            continue
        out.append(dict(o, dni=(teraz - ost).days if ost else None))
    out.sort(key=lambda o: (o["dni"] is not None, -(o["dni"] or 0)))
    return out


def wiersz(o):
    data = lambda d: d.strftime("%Y-%m-%d") if d else "nigdy"
    return [o["nazwa"], o["opis"], data(o["ostatnie"]),
            "nigdy" if o["dni"] is None else o["dni"],
            data(o["utworzone"]), "wylaczone" if o["wylaczone"] else "aktywne", o["ou"]]


def zapisz_xlsx(wyniki, sciezka):
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for rodzaj, lista in wyniki.items():
        ws = wb.create_sheet(rodzaj)
        ws.append(KOLUMNY[rodzaj])
        for c in ws[1]:
            c.font = Font(bold=True)
        for o in lista:
            w = wiersz(o)
            w[2], w[4] = o["ostatnie"], o["utworzone"]  # w Excelu jako daty
            ws.append(w)
        for kol, szer in zip("ABCDEFG", (22, 32, 18, 18, 14, 11, 45)):
            ws.column_dimensions[kol].width = szer
        for c in ws["C"][1:] + ws["E"][1:]:
            c.number_format = "yyyy-mm-dd"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(sciezka)


def run(domena, dni, out_dir, log, login="", haslo="", wylaczone=False, zrodlo=None):
    """Pobiera dane z AD, wybiera nieaktywne, zapisuje XLSX. Zwraca wyniki."""
    if dni < 14:
        log("UWAGA: lastLogonTimestamp jest replikowany z opoznieniem do 14 dni - "
            "wynik ponizej 14 dni jest niedokladny.")
    log("Odczyt AD: %s ..." % domena)
    dane = (zrodlo or pobierz)(domena, login, haslo)
    teraz = datetime.datetime.now()
    wyniki = {r: nieaktywne(dane[r], dni, teraz, wylaczone) for r in FILTRY}
    for r in FILTRY:
        log("%s: %d w AD, %d bez logowania od %d dni" % (r, len(dane[r]), len(wyniki[r]), dni))
    os.makedirs(out_dir, exist_ok=True)
    plik = os.path.join(out_dir, "nieaktywne_%s_%s.xlsx" % (domena, teraz.strftime("%Y-%m-%d")))
    zapisz_xlsx(wyniki, plik)
    log("Zapisano: " + plik)
    return wyniki

# ---------------------------------------------------------------- GUI ----


def gui():
    import tkinter as tk
    from tkinter import ttk, scrolledtext

    root = tk.Tk()
    root.title("Audyt AD - nieaktywne konta i komputery")
    root.geometry("1000x640")
    pad = dict(padx=6, pady=3)

    v_dom = tk.StringVar(value=os.environ.get("USERDNSDOMAIN", "").lower())
    v_dni = tk.StringVar(value="90")
    v_login = tk.StringVar()
    v_haslo = tk.StringVar()
    v_wyl = tk.BooleanVar(value=False)
    v_out = tk.StringVar(value=os.path.join(APP_DIR, "OUTPUT"))

    f = ttk.Frame(root)
    f.pack(fill="x", **pad)
    ttk.Label(f, text="Domena:").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_dom, width=30).grid(row=0, column=1, sticky="w", **pad)
    ttk.Label(f, text="Bez logowania od (dni):").grid(row=0, column=2, sticky="w", **pad)
    ttk.Spinbox(f, textvariable=v_dni, from_=14, to=3650, increment=30, width=7).grid(
        row=0, column=3, sticky="w", **pad)
    ttk.Checkbutton(f, text="Pokaz tez wylaczone", variable=v_wyl).grid(
        row=0, column=4, sticky="w", **pad)
    ttk.Label(f, text="Login (opcjonalnie):").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_login, width=30).grid(row=1, column=1, sticky="w", **pad)
    ttk.Label(f, text="Haslo:").grid(row=1, column=2, sticky="e", **pad)
    ttk.Entry(f, textvariable=v_haslo, width=20, show="*").grid(row=1, column=3, columnspan=2,
                                                                sticky="w", **pad)
    ttk.Label(f, text="puste = jako zalogowany uzytkownik Windows; inaczej np. URZAD\\admin",
              foreground="gray").grid(row=2, column=1, columnspan=4, sticky="w", padx=6)
    ttk.Label(f, text="Folder wyjsciowy:").grid(row=3, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_out, width=60).grid(row=3, column=1, columnspan=4,
                                                    sticky="w", **pad)

    btn = ttk.Button(root, text="Skanuj")
    btn.pack(pady=4)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, **pad)
    drzewa = {}
    for rodzaj, kolumny in KOLUMNY.items():
        ramka = ttk.Frame(nb)
        nb.add(ramka, text=rodzaj)
        t = ttk.Treeview(ramka, columns=kolumny, show="headings")
        for k, szer in zip(kolumny, (130, 200, 120, 110, 90, 80, 260)):
            t.heading(k, text=k)
            t.column(k, width=szer, anchor="w")
        sb = ttk.Scrollbar(ramka, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        drzewa[rodzaj] = t

    log_box = scrolledtext.ScrolledText(root, height=5)
    log_box.pack(fill="x", **pad)

    def log(msg):
        def put():
            log_box.insert("end", str(msg) + "\n")
            log_box.see("end")
        root.after(0, put)

    def pokaz(wyniki):
        for rodzaj, t in drzewa.items():
            t.delete(*t.get_children())
            for o in wyniki[rodzaj]:
                t.insert("", "end", values=wiersz(o))
            nb.tab(t.master, text="%s (%d)" % (rodzaj, len(wyniki[rodzaj])))

    def start():
        log_box.delete("1.0", "end")
        dom = v_dom.get().strip()
        if not dom or "." not in dom:
            return log("Podaj nazwe domeny, np. urzad.local")
        try:
            dni = int(v_dni.get())
        except ValueError:
            return log("Liczba dni musi byc liczba calkowita.")
        btn.config(state="disabled")

        def work():
            try:
                wyniki = run(dom, dni, v_out.get().strip('" '), log, v_login.get().strip(),
                             v_haslo.get(), v_wyl.get())
                root.after(0, lambda: pokaz(wyniki))
            except Exception as e:  # com_error, brak sieci, bledne haslo
                info = getattr(e, "excepinfo", None)
                e = info[2].strip() if info and info[2] else e
                log("BLAD: nie udalo sie odczytac AD (%s). Czy komputer widzi kontroler "
                    "domeny, a login i haslo sa poprawne?" % (e,))
            finally:
                root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    btn.config(command=start)
    if "--selftest" in sys.argv:
        root.after(200, root.destroy)
    import aktualizacja
    aktualizacja.start(root, "DawidBochno/Audyt-AD", "main", "audyt_ad.py")
    root.mainloop()

# ------------------------------------------------------------- selftest ----


def selftest():
    import tempfile
    import openpyxl

    class LI:  # IADsLargeInteger z ADO
        def __init__(self, n):
            self.HighPart, self.LowPart = n >> 32, n & 0xFFFFFFFF
            if self.LowPart >= 2 ** 31:  # COM oddaje LowPart ze znakiem
                self.LowPart -= 2 ** 32

    def ft(d):
        return int(d.timestamp() * 1e7) + FILETIME_1970

    teraz = datetime.datetime.now().replace(microsecond=0)
    dawno = teraz - datetime.timedelta(days=200)
    assert duza_liczba(LI(ft(dawno))) == ft(dawno)
    assert filetime(ft(dawno)) == dawno
    assert filetime(0) is None and filetime(0x7FFFFFFFFFFFFFFF) is None
    assert jednostka(r"CN=Kowalski\, Jan,OU=Kadry,OU=Urzad,DC=urzad,DC=local") == "Urzad/Kadry"
    assert jednostka("CN=PC01,CN=Computers,DC=urzad,DC=local") == "Computers"
    assert baza("urzad.gmina.local") == "DC=urzad,DC=gmina,DC=local"
    utc = datetime.datetime(2026, 1, 1, 12, 0)
    assert lokalny(utc) == lokalny(utc.replace(tzinfo=datetime.timezone.utc))

    def ad(nazwa, ost, utw, uac=512, opis=("Jan Testowy",), rodzaj="Konta"):
        return obiekt(rodzaj, {
            "sAMAccountName": nazwa, "name": nazwa, "displayName": opis[0],
            "operatingSystem": "Windows 11 Pro",
            "lastLogonTimestamp": LI(ft(ost)) if ost else None,
            "whenCreated": utw, "userAccountControl": uac,
            "distinguishedName": "CN=%s,OU=Kadry,DC=urzad,DC=local" % nazwa})

    d = lambda n: teraz - datetime.timedelta(days=n)
    utc_d = lambda n: d(n).astimezone(datetime.timezone.utc)
    konta = [ad("aktywny", d(3), utc_d(900)), ad("stary", d(200), utc_d(900)),
             ad("nigdy", None, utc_d(400)), ad("nowy", None, utc_d(5)),
             ad("wylaczony", d(300), utc_d(900), uac=514), ad("granica", d(91), utc_d(900))]
    komputery = [ad("PC01", d(120), utc_d(900), uac=4096, rodzaj="Komputery"),
                 ad("PC02", d(1), utc_d(900), uac=4096, rodzaj="Komputery")]
    assert konta[0]["ou"] == "Kadry" and konta[4]["wylaczone"] and not konta[0]["wylaczone"]
    assert komputery[0]["opis"] == "Windows 11 Pro"

    w = nieaktywne(konta, 90, teraz)
    assert [o["nazwa"] for o in w] == ["nigdy", "stary", "granica"], w
    assert w[1]["dni"] == 200 and w[0]["dni"] is None
    w = nieaktywne(konta, 90, teraz, wylaczone=True)
    assert [o["nazwa"] for o in w] == ["nigdy", "wylaczony", "stary", "granica"], w
    assert wiersz(w[0])[2:4] == ["nigdy", "nigdy"] and wiersz(w[1])[5] == "wylaczone"

    tmp = tempfile.mkdtemp()
    lines = []
    wyn = run("urzad.local", 7, tmp, lines.append,
              zrodlo=lambda dom, l, h: {"Konta": konta, "Komputery": komputery})
    assert all(ord(ch) < 128 for line in lines for ch in line), lines
    assert any("UWAGA" in x for x in lines) and [o["nazwa"] for o in wyn["Komputery"]] == ["PC01"]
    assert any("Konta: 6 w AD, 3 bez logowania od 7 dni" in x for x in lines), lines
    plik = [f for f in os.listdir(tmp) if f.endswith(".xlsx")]
    assert plik and plik[0].startswith("nieaktywne_urzad.local_"), plik
    wb = openpyxl.load_workbook(os.path.join(tmp, plik[0]))
    assert wb.sheetnames == ["Konta", "Komputery"]
    ws = wb["Konta"]
    assert [c.value for c in ws[1]] == KOLUMNY["Konta"]
    assert ws["A2"].value == "nigdy" and ws["C2"].value is None and ws["D2"].value == "nigdy"
    assert ws["C3"].value == d(200) and ws["D3"].value == 200 and ws.auto_filter.ref == "A1:G4"
    assert wb["Komputery"]["A2"].value == "PC01" and wb["Komputery"]["B2"].value == "Windows 11 Pro"

    import aktualizacja
    aktualizacja.selftest()
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        if "--gui" in sys.argv:
            gui()
    else:
        gui()
