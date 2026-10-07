# Audyt AD — instrukcja dla administratora

Ta instrukcja prowadzi przez cały przegląd domeny krok po kroku: od
pierwszego skanu, przez czytanie wyników, po naprawę znalezionych problemów
i przegląd okresowy. Opis programu i jego kontroli jest w [README](../README.md).

## Spis treści

1. [Zanim zaczniesz](#1-zanim-zaczniesz)
2. [Pierwszy skan](#2-pierwszy-skan)
3. [Jak czytać wyniki](#3-jak-czytać-wyniki)
4. [Co zrobić z każdą uwagą](#4-co-zrobić-z-każdą-uwagą)
5. [Przegląd okresowy](#5-przegląd-okresowy-procedura)
6. [Rozwiązywanie problemów](#6-rozwiązywanie-problemów)
7. [Słownik](#7-słownik)

---

## 1. Zanim zaczniesz

| Potrzebne | Uwagi |
|-----------|-------|
| Komputer w domenie **albo** komputer w tej samej sieci co kontroler domeny | z domu przez VPN też działa, jeśli VPN daje dostęp do kontrolera |
| Zwykłe konto domenowe | do odczytu nie trzeba być administratorem domeny |
| Python 3.9+ i jednorazowe `install.bat` | patrz [README → Instalacja](../README.md#instalacja-jednorazowo) |
| Excel albo LibreOffice | do otwarcia raportu (opcjonalnie) |

**Program niczego nie zmienia w domenie.** Polecenia naprawcze w rozdziale 4
uruchamiasz sam, świadomie, po sprawdzeniu każdego konta. Wymagają one
modułu PowerShell **ActiveDirectory** (RSAT) i uprawnień administratora:

```powershell
# Windows 10/11 - instalacja RSAT dla AD (PowerShell jako administrator)
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

Na kontrolerze domeny moduł jest już zainstalowany.

---

## 2. Pierwszy skan

1. Uruchom **`uruchom.bat`**.
2. Sprawdź pole **Domena**. Na komputerze w domenie jest wypełnione
   automatycznie (np. `urzad.local`). Jeśli jest puste, wpisz pełną nazwę
   DNS domeny. Znajdziesz ją w *Ustawienia → System → Informacje* albo
   w wyniku polecenia `echo %USERDNSDOMAIN%` w wierszu poleceń.
3. Zostaw **Login** i **Hasło** puste, jeśli komputer jest w domenie.
   Na komputerze spoza domeny wpisz `DOMENA\login` albo `login@domena.local`.
4. Ustaw progi. Na pierwszy raz domyślne wystarczą:
   - **Bez logowania od (dni)** = 90,
   - **Hasło starsze niż (dni)** = 365.
5. Kliknij **Skanuj**. Na dole okna pojawi się przebieg, np.:

   ```
   Odczyt AD: urzad.local ...
   Nieaktywne konta: 12
   Nieaktywne komputery: 7
   Uprzywilejowani: 6
   Ryzyka kont: 9
   Ryzyka komputerow: 15
   Podsumowanie: 6 kontroli z UWAGA
   Zapisano: C:\Programy\Audyt AD\OUTPUT\audyt_AD_urzad.local_2026-10-07.xlsx
   ```

6. **Sprawdź wynik na kilku kontach**, które znasz. Otwórz
   *Użytkownicy i komputery usługi Active Directory* (`dsa.msc`) → *Widok →
   Funkcje zaawansowane* → konto → *Edytor atrybutów* → `lastLogonTimestamp`.
   Data powinna się zgadzać z kolumną **Ostatnie logowanie** (z dokładnością
   do 14 dni, patrz [README](../README.md#dokładność-daty-logowania)).

---

## 3. Jak czytać wyniki

Zacznij od zakładki **Podsumowanie**. To lista kontrolna całej domeny.

- **UWAGA** (czerwony wiersz) — coś wymaga decyzji. Nie zawsze błąd, ale
  zawsze coś, co audytor zapyta.
- **OK** — kontrola przeszła.
- **INFO** — informacja do wiadomości (np. liczba obiektów).

Potem przejdź po zakładkach w tej kolejności (od największego ryzyka):

1. **Uprzywilejowani** — czy każdą osobę na liście znasz i czy nadal
   potrzebuje uprawnień administratora? Każde konto administratora, którego
   nie potrafisz wyjaśnić, to priorytet.
2. **Ryzyka kont** — szczególnie *bez preautoryzacji Kerberos*,
   *hasło niewymagane* i *ma SPN* z bardzo starym hasłem.
3. **Nieaktywne konta** — do wyłączenia po potwierdzeniu w kadrach.
4. **Ryzyka komputerów** — systemy bez wsparcia, LAPS.
5. **Nieaktywne komputery** — do wyłączenia, później usunięcia.

Pomocne w oknie:

- **Szukaj** — wpisz nazwisko, wydział (`Kadry`) albo rodzaj uwagi
  (`LAPS`, `SPN`); filtruje wszystkie zakładki naraz.
- **Kliknięcie nagłówka** — sortowanie (np. po *Dni bez logowania*).
- **Dwuklik na wierszu** — całe uwagi w okienku.

W pliku Excela te same zakładki mają filtry w nagłówkach — wygodne do
przekazania listy kadrom albo przełożonemu.

---

## 4. Co zrobić z każdą uwagą

> **Zasada:** najpierw *wyłącz*, poczekaj (np. 30 dni), dopiero potem *usuń*.
> Wyłączenie da się cofnąć jednym kliknięciem; usunięcia (bez włączonego
> Kosza AD) nie. Przed każdą zmianą upewnij się, że konto nie jest używane
> przez usługę, drukarkę, skaner albo aplikację.

W poleceniach podmień `jkowalska`, `UG-PC-017` itd. na nazwy z raportu.

### Nieaktywne konta

1. Wyślij listę do kadr: czy te osoby nadal pracują (urlop
   macierzyński/wychowawczy, długie L4 — to częste i poprawne powody)?
2. Konta osób, które odeszły — wyłącz i opisz:

   ```powershell
   Disable-ADAccount -Identity jkowalska
   Set-ADUser -Identity jkowalska -Description "Wylaczone 2026-10-07 - audyt AD, brak logowania 412 dni"
   ```

3. Po okresie przejściowym (np. 30–90 dni, zgodnie z polityką urzędu)
   usuń konto, a wcześniej zabezpiecz dane (skrzynka, pliki domowe).

Konto nigdy nie logowane (*Ostatnie logowanie = nigdy*) zwykle zostało
założone „na zapas” albo dla osoby, która nie podjęła pracy — wyłącz.

### Nieaktywne komputery

```powershell
Disable-ADAccount -Identity 'UG-PC-017$'     # znak $ na końcu - konto komputera
```

Komputer, który wróci do sieci (np. laptop po długim urlopie), po
wyłączeniu konta przestanie uwierzytelniać się w domenie — wtedy włącz go z powrotem
(`Enable-ADAccount`). Sprawdź też, czy w AD nie ma kluczy BitLockera,
których będziesz potrzebować, zanim usuniesz obiekt komputera.

### Uprzywilejowani

| Uwaga | Działanie |
|-------|-----------|
| osoba, która nie powinna być adminem | usuń z grupy: `Remove-ADGroupMember -Identity "Administratorzy domeny" -Members jkowalska` |
| nieaktywne od ponad N dni | wyłącz konto administratora albo usuń z grupy |
| adminCount=1 bez członkostwa (pozostałość) | wyczyść atrybut i przywróć dziedziczenie uprawnień — patrz niżej |
| codzienna praca na koncie administratora | załóż osobne konto do pracy biurowej (bez uprawnień) i osobne `adm.…` do administracji |

**Czyszczenie pozostałości `adminCount`:**

```powershell
Set-ADUser -Identity bzielinska -Clear adminCount
```

Następnie w `dsa.msc` → konto → *Zabezpieczenia* → *Zaawansowane* →
**Włącz dziedziczenie**. Bez tego konto zachowuje zablokowane uprawnienia
(np. helpdesk nie zresetuje mu hasła).

### Ryzyka kont

| Uwaga | Działanie | Polecenie |
|-------|-----------|-----------|
| hasło niewymagane | wyłącz flagę, nadaj hasło | `Set-ADUser jkowalska -PasswordNotRequired $false` |
| hasło z odwracalnym szyfrowaniem | wyłącz flagę, **potem zmień hasło** (stare zostaje zapisane odwracalnie do zmiany) | `Set-ADUser jkowalska -AllowReversiblePasswordEncryption $false` |
| hasło nigdy nie wygasa | dla zwykłych osób wyłącz; dla kont technicznych — długie losowe hasło (25+ znaków) albo konto gMSA | `Set-ADUser jkowalska -PasswordNeverExpires $false` |
| delegacja nieograniczona | zamień na delegację ograniczoną (zakładka *Delegowanie* w `dsa.msc`); wymaga sprawdzenia, której usłudze była potrzebna | `Set-ADAccountControl jkowalska -TrustedForDelegation $false` |
| tylko szyfrowanie DES | wyłącz flagę, zmień hasło | `Set-ADAccountControl jkowalska -UseDESKeyOnly $false` |
| bez preautoryzacji Kerberos | wyłącz flagę — wyjątkowo rzadko potrzebna (stare systemy uniksowe) | `Set-ADAccountControl jkowalska -DoesNotRequirePreAuth $false` |
| ma SPN (Kerberoasting) | hasło konta usługi: 25+ losowych znaków, zmienione; najlepiej przejść na gMSA; usunąć niepotrzebne SPN | `Get-ADUser svc_backup -Properties servicePrincipalName` |
| hasło niezmieniane od N dni | wymusić zmianę przy następnym logowaniu | `Set-ADUser jkowalska -ChangePasswordAtLogon $true` |
| konto wygasło | wyłączyć / usunąć zgodnie z procedurą jak dla nieaktywnych | |

Kont technicznych, które **świadomie** mają wyjątek (np. konto skanera
z hasłem bez wygasania, ale długim i losowym), nie trzeba oglądać przy
każdym audycie — dopisz je do **Wykluczeń** (przycisk *Wykluczenia…*),
najlepiej z komentarzem, kto i dlaczego zdecydował:

```
# skaner na parterze - haslo 30 znakow, zmiana raz w roku (J.Kowalski, 2026-10)
skaner
```

### Ryzyka komputerów

| Uwaga | Działanie |
|-------|-----------|
| system bez wsparcia | zaplanuj aktualizację lub wymianę; do tego czasu odizoluj komputer (osobny VLAN, brak dostępu do internetu). Windows 10 z wykupionym ESU — odnotuj to w dokumentacji i dodaj do wykluczeń. |
| koniec wsparcia (za N dni) | zaktualizuj do nowszej wersji (np. Windows 11 → *Windows Update → najnowsza wersja*) |
| brak LAPS | sprawdź, czy komputer jest w OU objętej zasadą LAPS i czy dostaje GPO (`gpresult /r` na komputerze) |
| LAPS: hasło nieodnowione | LAPS na komputerze nie działa: sprawdź dziennik *Microsoft → Windows → LAPS* w Podglądzie zdarzeń, wymuś `Invoke-LapsPolicyProcessing` |
| delegacja nieograniczona | jak dla kont — zamień na ograniczoną |

### Podsumowanie

| Kontrola z UWAGĄ | Działanie |
|------------------|-----------|
| Minimalna długość hasła < 12 | GPO *Default Domain Policy* → *Zasady konta → Zasady haseł* albo `Set-ADDefaultDomainPasswordPolicy -Identity urzad.local -MinPasswordLength 12`. Zmiana dotyczy haseł ustawianych **od teraz**. Uprzedź pracowników. |
| Brak blokady konta | `Set-ADDefaultDomainPasswordPolicy -Identity urzad.local -LockoutThreshold 10 -LockoutDuration 00:15:00 -LockoutObservationWindow 00:15:00` |
| MachineAccountQuota > 0 | `Set-ADDomain -Identity urzad.local -Replace @{"ms-DS-MachineAccountQuota"="0"}`. Komputery dołącza wtedy do domeny tylko administrator (albo konto z delegowanym uprawnieniem w OU). |
| Hasło krbtgt starsze niż 180 dni | zmień hasło **dwa razy, w odstępie co najmniej 10 godzin** (czas ważności biletów Kerberos). Użyj skryptu Microsoft `New-KrbtgtKeys.ps1` — sprawdza replikację przed zmianą. **Nie zmieniaj dwa razy pod rząd** — wszyscy zostaną wylogowani z usług domenowych. |
| Konto Gość włączone | `Disable-ADAccount -Identity Gość` (nazwa zależy od języka: *Gość* / *Guest*) |
| Administrator — stare hasło | zmień hasło na długie, losowe, zapisane w sejfie haseł / kopercie; do codziennej administracji używaj imiennych kont `adm.…` |
| LAPS nie wdrożony | wdrożyć **Windows LAPS** (wbudowany w Windows 10/11 i Server 2019+ od aktualizacji z kwietnia 2023): `Update-LapsADSchema`, `Set-LapsADComputerSelfPermission -Identity "OU=Komputery,DC=urzad,DC=local"`, GPO *System → LAPS*. Szczegóły: dokumentacja Microsoft „Windows LAPS”. |
| Poziom funkcjonalny domeny starszy niż 2012 R2 | najpierw wymień stare kontrolery domeny, potem podnieś poziom (*Domeny i relacje zaufania AD*) — operacja nieodwracalna, wymaga planu |

Po wprowadzeniu zmian uruchom **Skanuj** ponownie i porównaj raporty.

---

## 5. Przegląd okresowy (procedura)

Proponowany rytm: **co kwartał** pełny przegląd, **co miesiąc** szybki rzut
oka na *Uprzywilejowani* i *Podsumowanie*.

Lista kontrolna kwartalnego przeglądu:

- [ ] Skan z progami 90 / 365 dni, raport XLSX zapisany w dokumentacji IT
      (np. `\\serwer\IT\Audyty AD\2026-Q4\`).
- [ ] *Uprzywilejowani* — każda osoba potwierdzona przez przełożonego;
      zmiany odnotowane.
- [ ] *Nieaktywne konta* — lista wysłana do kadr, odpowiedź zarchiwizowana,
      konta wyłączone.
- [ ] Konta wyłączone w poprzednim kwartale — usunięte (po zabezpieczeniu danych).
- [ ] *Ryzyka kont* — nowe pozycje wyjaśnione albo dodane do wykluczeń
      z uzasadnieniem.
- [ ] *Ryzyka komputerów* — systemy bez wsparcia: plan wymiany z terminem.
- [ ] *Podsumowanie* — liczba UWAG nie wzrosła względem poprzedniego kwartału.

Raport z podpisem osoby odpowiedzialnej to dowód okresowego przeglądu
uprawnień — wymaganego m.in. przez RODO (art. 32 ust. 1 lit. d — regularne
testowanie i ocena skuteczności środków) i rozporządzenie KRI (§ 20).

**Uwaga:** raport zawiera dane osobowe i opis słabych punktów domeny.
Przechowuj go w miejscu z ograniczonym dostępem i nie wysyłaj otwartą
pocztą.

---

## 6. Rozwiązywanie problemów

| Objaw (komunikat na dole okna) | Przyczyna | Rozwiązanie |
|--------------------------------|-----------|-------------|
| `Podaj nazwe domeny, np. urzad.local` | puste pole Domena (komputer spoza domeny) | wpisz pełną nazwę DNS domeny |
| `BLAD: ... (Serwer nie działa.)` | komputer nie widzi kontrolera domeny | sprawdź sieć/VPN; `nltest /dsgetdc:urzad.local` w wierszu poleceń musi zwrócić kontroler; sprawdź, czy DNS komputera wskazuje na kontroler |
| `BLAD: ... (Nazwa użytkownika lub hasło jest niepoprawne.)` | zły login/hasło | login w formacie `DOMENA\login` lub `login@domena.local`; sprawdź Caps Lock |
| `BLAD: ... (Domena określona nie istnieje ...)` | literówka w nazwie domeny albo komputer spoza domeny bez loginu | popraw nazwę; poza domeną wpisz login i hasło |
| `OU podaj jako pelna sciezke ...` | pole OU w złym formacie | skopiuj *distinguishedName* jednostki z `dsa.msc` → *Edytor atrybutów*, np. `OU=Urzad,DC=urzad,DC=local` |
| `BLAD: nie mozna zapisac ... - zamknij ten plik w Excelu.` | raport z dzisiaj jest otwarty w Excelu | zamknij plik i skanuj ponownie |
| `UWAGA: lastLogonTimestamp jest replikowany z opoznieniem ...` | próg poniżej 14 dni | to tylko ostrzeżenie; dla krótkich progów wynik jest przybliżony |
| okno się nie otwiera po `uruchom.bat` | Python bez tkinter albo nie zainstalowano bibliotek | uruchom `install.bat` i przeczytaj komunikat; przy braku tkinter zainstaluj Pythona ponownie z opcją „tcl/tk and IDLE” |
| `install.bat`: błąd instalacji bibliotek | brak internetu / UTM blokuje `pypi.org` | poproś o odblokowanie `pypi.org` i `files.pythonhosted.org` albo zainstaluj biblioteki z innego komputera |
| konto, które na pewno działa, jest „nieaktywne” | logowanie tylko do poczty w chmurze albo aplikacji z własnymi kontami; albo opóźnienie replikacji | sprawdź `lastLogon` na każdym kontrolerze (`Get-ADUser jkowalska -Properties lastLogon -Server DC01`) |
| zakładka *Uprzywilejowani* pusta | konto, którym skanujesz, nie widzi grup (rzadkie, zmienione uprawnienia odczytu) | skanuj kontem z domyślnymi uprawnieniami odczytu albo kontem administratora |
| brak wierszy LAPS w Podsumowaniu mimo wdrożonego LAPS | hasła LAPS przechowywane tylko w Entra ID (chmura) | program widzi tylko LAPS zapisywany w AD |

Jeśli komunikat jest inny, skopiuj go (zaznacz w oknie programu, `Ctrl+C`)
i zgłoś w [Issues](https://github.com/DawidBochno/Audyt-AD/issues)
**bez danych z raportu** (bez nazwisk i nazw kont).

---

## 7. Słownik

| Pojęcie | Znaczenie |
|---------|-----------|
| **AD (Active Directory)** | katalog kont użytkowników, komputerów i grup w sieci Windows |
| **Kontroler domeny (DC)** | serwer, na którym działa AD |
| **OU (jednostka organizacyjna)** | „folder” w AD, np. *Urzad/Kadry* |
| **LAPS** | mechanizm Microsoft, który na każdym komputerze ustawia inne, losowe, regularnie zmieniane hasło lokalnego administratora i zapisuje je w AD |
| **krbtgt** | konto systemowe, którego klucz podpisuje wszystkie bilety logowania w domenie |
| **Kerberoasting** | atak: każdy użytkownik domeny może pobrać zaszyfrowany bilet do usługi (konto z SPN) i łamać hasło tej usługi offline |
| **AS-REP roasting** | podobny atak na konta bez preautoryzacji Kerberos — nie trzeba nawet mieć konta w domenie |
| **SPN** | „adres” usługi w Kerberos (np. `MSSQLSvc/sql01`); konto użytkownika z SPN to konto usługi |
| **Delegacja nieograniczona** | serwer może działać w imieniu każdego użytkownika, który się do niego zaloguje — także administratora domeny |
| **adminCount / AdminSDHolder** | mechanizm AD chroniący konta w grupach administracyjnych; po usunięciu z grupy znacznik zostaje |
| **MachineAccountQuota** | ile komputerów może dodać do domeny zwykły użytkownik (domyślnie 10) |
| **gMSA** | konto usługi zarządzane przez AD — hasło 240 znaków zmieniane automatycznie |
| **lastLogonTimestamp** | data ostatniego logowania kopiowana między kontrolerami (z opóźnieniem do 14 dni) |
| **ESU** | płatne rozszerzone aktualizacje zabezpieczeń dla systemów po końcu wsparcia |
