Attribute VB_Name = "VersionModule"
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.



' 3.98
' dodatkowe warunki przez co mamy dwie formy pracy - design mode i production mode
' QT mod - nieco zmieniony layout pivotow - do zaprezentowania.
' QT mod 2 - przygotowanie bezposrednio pod 6p zrzut - osobny modul narazie pusty bedzie guzik dodoatkowy w ribbon

' wizard 3.97 - 2016-03-21 to be shared

' dodatkowy guzik generowania unique ID
'
'
''
' unique ID jest spore zatem zmniejszylem wielkosc fontu do: 2
'
' validation rozszerzone przy dodawaniu formul rowniez sie dodaje logika nadawania obostrzenia
' na delivery reconf oraz oncost conf
''
'
'
''
' upada pomysl z dynamicznymi alertami podczas dodawania PUSow
''
'
'
' ===================================================
' wizard 3.96 dev - to be implemented 2016-03-15
' kopia zapasowa - repo - rozszerzenie details
' dodatkowy guzik generowania unique ID
'
'
''
' unique ID jest spore zatem zmniejszylem wielkosc fontu do: 2
'
' validation rozszerzone przy dodawaniu formul rowniez sie dodaje logika nadawania obostrzenia
' na delivery reconf oraz oncost conf
''
'
'
' ===================================================
' wizard 3.95 dev - to be implemented 2016-03-15
' kopia zapasowa - repo - rozszerzenie details
' ===================================================
' wizard 3.94 dev - to be implemented 2016-03-14
' ===================================================
'
'
''
' twarde alerty podczas dodawania pusow
' plus ewentualne doapsowanie nowej daty del conf w przypadku gdy edytowany lub dodawany jest pus
'
' dodatkowo validation jako rozwiazanie dla kolumn del reconf i oncost conf nie powiodlo sie
' musze znalezc inne rozwiazanie - to jeszcze wisi
'
'
' ID parametr w arkuszu details - handler generator uniqa
'
' ===================================================

' wizard 3.93 dev - to be implemented 2016-03-11
' ===================================================
'
'
''
' masowe kopiowanie pono # lub customow w ogole
' rozwiazanie to zapewne bedzie dynamicznie zmieniac
' mozliwosc wyboru w listbox'ie z pojedynczego na wielokrotny
''
'
'
''
' dodatkowo rozszerzenie delivery confirmation
''
'
'
' ===================================================



' wizard 3.92 dev - to be implemented 25,02,2016
' ===================================================
' rozszerzyc mozliwosci podmiany niestd
' 1. dla wielu kolumn na raz gdy mamy do przekopiowania customy np pono numery
' quarter time issue - w supp nm pojawia sie duns! !!!


''
' fajnie by bylo jeszcze zrobic masowe kopiowanie
''

''
' w tym ponizej uzylem std rozwiazania excela! 2016-03-09
' reconfirmation moze byc albo OK albo NOK nic wiecej
''


''
' poprawka w QTHandler class - w sumie prosta sprawa nalezalo zmieniac przypisanie juz gotowej
' wlasciwosci klasy do odpowiedniej komorki (wczesniej odruchowo dwa razy wpisalem .duns)
''



' wizard 3.91
' ===================================================
' przygotowanie pod nowe del confy
' troche zmian - sprawdzenie jeszcze raz logiki dzialania przekopiowywania danych

' wizard 3.87
' ===================================================
' dodane pole w details  platform (przed account)
' transportation logika dynamiczna dodania danych
' cos nie tak jest z wizardami details (szczegolnie w toggle'ach)
' zatem silowe przypisanie dodatkowe
' jednak aktualnie wersja udostepniona to 3.86
' 3.87 pewnie jest przejsciowe i nigdy nie zostanie wdrozone
' a jednak 3.87 dodatkowo dalem poprawiona logike dat - co wadliwie
' sie nie przesuwaly - dodalem do internetowej formuly inna
' implementacje - okazalo sie ze pierwotna jest beznadziejna
' zatem wykorzystalem na glupa prosta petle ktora leci od
' pierwszego stycznia i dzien po dniu dopoki nie wpadnie
' pierwszy dzien pozadanego tygodnia kalendarzowego - to jako tako rozwiazuje problem
' ===================================================

' wizard 3.86
' ===================================================
' dodane pole w details  platform (przed account)
' transportation logika dynamiczna dodania danych
' ===================================================

' wizard 3.85
' ===================================================

' 1. nowy details nowy form plus rozszerzony stary
' 2. postawienie flag msgbox ze dopisanie danych (funkcjonalnosc) nie  jest jeszcze
' gotowa
' 3. zastanawiam sie nad dorzuceniem dodatkowego formularza wypelnienia arkusza details
' jednoczesnie mielibysmy ladny wglad na cale dane
' 4. pamiec podreczna dla account (pole details



' wizard 3.84
' ===================================================

' edycja del conf z poziomu formularzy pus
' dodanie elementu mrd data w arkusz details


' wizard 3.83
' ===================================================
' del conf w edit pus - kolorowanie dynamiczne
' dodanie zera gdy < 10 dla YxxxxCWxx OK juz!

' w details format yyyy-mm-dd hh:mm niech zostanie
' i niech collector zajmie sie tym na wlasna reke

' edycja del conf z poziomu formularzy pus
' wciaz to be done


' wizard 3.82
' ===================================================
' del conf w edit pus + male fixy

' wizard 3.81
' ===================================================
' del conf w edit pus

' wizard 3.80
' ===================================================
' dodanie ecycji rowniez dla synchro
' brakuje dat dla dodawania

' wizard 3.79
' ===================================================
' synchro z jenny - kod pickups handler nie jest skonczony
' znaczaco przerobiony
' inner_logic_for_duns_page_on_active_master_worksheet()
' wydzielenie dodatkowych subow:
'
'
' normal_init_for_dodaj_pus
' jenny_init_for_dodaj_pus


' wizard 3.78
' ===================================================
' adjust on quarter time

' wizard 3.77
' ===================================================
' pobieranie danych z pliku external - calkowita podmiana gotowa
' dodatkowo opcja synchro komentarzy


' wizard 3.76
' ===================================================
' arkusz pusow w bardziej dogodnej wersji


' wizard 3.75
' ===================================================
' wiecej pivotow intransit


' wizard 3.74
' ===================================================
' Quarter first draft
' automatyczny pivot na new workbook

' wizard 3.73
' ===================================================
' Quarter first draft
' dodatnie klas oblugujacych QTHandler
' PUSBucket - jako komponent QTHandler

' wizard 3.72
' ===================================================
' dodanie do listy def conf danych
' jeszcze mala poprawka z add date wizard (zeby na poczatku bylo today)
' brak implementacji pod Quartera jeszcze
' poprawiony delivery confirmation
' ===================================================

' wizard 3.71
' ===================================================
' funkcjonalnosc usun pusa cos z nia nie tak
' zmienie pewnie delte entire row na clear
' + blad przy delivery confirmation



' wizard 3.70
' ===================================================
' rozszerzenie ograniczenia do 8 osob
' podmiana prawidolowo wyswietlanego info z msgboxa
' jeden hack jest - to jest kontrola uzytkonikow mocno zagniezdzona
' w prywatnym subie ustawiania flagi na koncu w sub dodaj dla aktywnej
' zakladki duns sub:
' ustaw_tylko_na_pierwsze_puste_miejsce
' wizard 3.69
' ===================================================
' rozszerzenie ograniczenia do 8 osob
' podmiana prawidolowo wyswietlanego info z msgboxa
' jeden hack jest - to jest kontrola uzytkonikow mocno zagniezdzona
' w prywatnym subie ustawiania flagi na koncu w sub dodaj dla aktywnej
' zakladki duns sub:
' ustaw_tylko_na_pierwsze_puste_miejsce
'
' ===================================================

' wizard 3.68
' ===================================================
'
' ostatnia wersja blok przy wiecej niz 4 osobach
' blad msgboxwej info
'
' ===================================================



' testy synchro na werjsach! 3.62 - 3.67 (duzo zmian zwiazanych z elementami pracy rownoleglej)
' usuniecie ThisWorkbook.Save

' wizard 3.61
' ===================================================
'
' 1. blokowanie przy DODAJ PICKUP tego samego PUS #!
' 2. wszystko po stronie edycji czyli:
' edycja delivery i pickup date
' edycja juz istniejacych pn
' mozliwosci dodania pn i ich usuwania
' transakcja szybki zapis wszystkiego
'
' ===================================================

' wizard 3.60
' ===================================================
'
'
' wszystko ok
' to jest wersja testowa pod pattern edycji pusow
'
' ===================================================




' wizard 3.59
' ===================================================
'
'
' zmiana formuly sumif (bardziej lokalna zeby dzialalo
' nawet gdy zaczniemy sortowac
' wczesniej nie hulalo
'
' dorzucenie selekcji przy selekcji na dodawaniu pusow
'
'
' sprawdzenie warunku chronologicznosci data del date i
' pickup date
'
'
' ===================================================
' wizard 3.58
' ===================================================
'
'
' niedzielna przygoda z dodawaniem mimo wszystko po pnie
' a tak poza tym:
' 2. formula handler - MIGHTY FORMULA REPAIRING! done!
'
'
' ===================================================

' wizard 3.57
' ===================================================
'
' zmieny dotyczace wszystkich scenariuszy
' 128 mozliwosci ustawienia ukladow formul statusu ok / nok
' powazne zmiany w logice dodawania pusow
' zblokowanie mozliwosci dodawania pusow po jednym pn'ie
' a tak poza tym:
' 2. formula handler - on hold
' 3. problemy z ustatecznym ustawieniem formul
'
'
' ===================================================


' wizard 3.56
' ===================================================
' 0. od tej wersji zero kontroli na starcie i przed zamknieciem
'   (za duzo zlego sie dzieje z plikami share)
' 1. pusy add / edit / delete
' 2. formula handler - on hold
' 3. zmiany amrd i bmrd zm. na nazwy mrd1 i mrd2
' 4. zmiana ukladu kolumn pod mrd1 i mrd2
' 5. guziki projekt normlany i podwojny projekt
' ===================================================

' wizard 3.55
' ===================================================
' 0. od tej wersji zero kontroli na starcie i przed zamknieciem
'   (za duzo zlego sie dzieje z plikami share)
' 1. optymalizacja kodu w logice dodawania pusow - edytowanie juz istniejacych PUSow + usuwanie
' 2. formula handler - on hold
' 3. 2015-10-05 - draft implementacji inner_logic_for_edit_puses_on_active_master_worksheet
' 4. dodawanie pusow weryfikacja ile zostalo do zrobienia pusow w zaleznosci od Orderow
' 5. wraz z powyzszym duze zmiany z nazewnictwem BMRD i AMRD - tutaj nie ma jeszcze klepniecia ze strony
' mgmt
'
' ===================================================

' wizard 3.54
' ===================================================
' 0. od tej wersji zero kontroli na starcie i przed zamknieciem
'   (za duzo zlego sie dzieje z plikami share)
' 1. optymalizacja kodu w logice dodawania pusow - edytowanie juz istniejacych PUSow
' 2. formula handler - on hold
' 3. 2015-10-05 - draft implementacji inner_logic_for_edit_puses_on_active_master_worksheet
'
' ===================================================

' wizard 3.53
' ===================================================
' 1. optymalizacja kodu w logice dodawania pusow - edytowanie juz istniejacych PUSow
' 2. formula handler - on hold
' 3. 2015-10-05 - draft implementacji inner_logic_for_edit_puses_on_active_master_worksheet
' 4. dalej sie cos zwiesza podczas proby uruchomienia plikow master
' 5. chyba musze sie pozbyc logiki na starcie i zamknieciu masterow
'
' ===================================================

' wizard 3.52
' ===================================================
' 1. optymalizacja kodu w logice dodawania pusow on hold
' 2. dodatkowe 2 guziki zrob backup (kod juz byl wczesniej) plus on off forms dla odpowiednich pol - jednak chyba bedzie trzeba je usunac
'   badz przynajmniej zmodyfikowac pod wzgledem tym aby delivery conf. form sie pojawial zawsze
' 3. formula handler - on hold
' 4. zmiana orderu kolumn
'
' ===================================================

' wizard 3.51
' ===================================================
' 1. optymalizacja kodu w logice dodawania pusow on hold
' 2. dodatkowe 2 guziki zrob backup (kod juz byl wczesniej) plus on off forms dla odpowiednich pol
' 3. formula handler - on hold
' 4. przygotowanie pod zmiany orderu kolumn
'
' ===================================================

' wizard 3.42
' ===================================================
' 1. optymalizacja kodu w logice dodawania pusow
' 2. dodatkowe 2 guziki zrob backup (kod juz byl wczesniej) plus on off forms dla odpowiednich pol
' 3. formula handler - on hold
' 4. dodane nowe kolumny z powodu zakleszczenia timingow w T/D dla przed MRD i po MRD
'
' ===================================================

' wizard 3.41
' ===================================================
' 1. optymalizacja kodu w logice dodawania pusow
' 2. pojawil sie problem zwiazany z toggleButton jesli uruchamiam excela od poczatku
' ustawienie filtracji zapisanej zostaje, jednak guziki juz nie wracaja do wczesniejszej ustawien
' wszystkie startuja od pozycji nie wcisnietej
' rozwiazanie: przy event open workbook dajemy swiezy filtr od poczatku
'
' ===================================================

' wizard 3.4
' ===================================================
' 1. fma i sel i fup code toggle button
'
' ===================================================

' wizard 3.3

' 1. system dodatkowej filtracji po selekcji zarowno pojedyncze
' jak i wielu aktywnych komorek jednej kolumny
' 2. sprawdzenie czy arkusz register naprawde jest tak bardzo potrzebny
'
' ===================================================

' wizard 3.2
' system wpisu delivery confirmation z listy
' ===================================================


' wizard 3.1
' master oddany do rak koordynatorow
' nowa kolumna first pickup date
' ===================================================


' wizard 2.12
' ===================================================
' historia krotka
' z dodawaniem nowych pickuppw
' dodana funkcja w pickups handler aby dodawalo zera z lewej strony
' na formualrze

' wizard 2.11
' ===================================================
' od tego wizarda jest jasne ze logika walidacji w masterze i pickupach odzielone sa od siebie
' w dwoch roznych klasach
' master posiada do tego wyodrebniony obiekt
'
' natomiast pusy wszystko zalatwiaja we wlasnym domu
' ===================================================


' wizard 2.0.xlsm_backup__timestamp_42268_3232986111
' ===================================================
' pierwsza baza na dzien 21 september 2015
' jakies dodatkowe
' oddanie gosi nowej wersji
'
' ARKUSZ HISORII do rozpatrzenia (albo bardziej logu)
' przygotowanie handlera formul




' ===================================================

' wizard 2.0.xlsm_backup__timestamp_42267_88
' ===================================================
' baza pod kolejne zmiany
' gotowy uklad wstawiania pusow
' uklad kolumn ok
' wizard pod details


' wizard 2.0.xlsm_backup__timestamp_42267_6275
' ===================================================

' baza pod przygotowanie kolejnych napraw bugow
' edycja i new details project
' ppap gate i pickup date zostaly juz odseparowane nazewniczo

' 1. dopisac trzeba jeszcze mozliwosc wyboru daty #NA!
' 2. faza lista wstepne przygotowanie ? (ale nie wydjae mi sie to najbardziej slusznym rozwiazaniem
' tutaj ewentualnie pomyslalem tylko o czesciowym patternie ktory poprawia wielkosc liter
' jesli pattern zostal czesciowo rozpoznany

' 3. czesto sie pojawia msgbox od adjusta - ktory nigdy nie powinien sie pokazac co jest troche meczace
' musze zoabczyc z czego wynika problem :)

' 4. wlasciwie filtered zaminilem do ograniczania wyswietlania jednego dunsu przy pomocy odpowiedniego checkboxa


' ===================================================

' wizard 2.0.xlsm_backup__timestamp_42267_5278472222
' ===================================================
' wersja bazowa dla zmian po pierwszym spotkaniu z koordynatorem

' poprawic nalezy blad edycji i dodawania projektu
' te bug do wspiania na liste na sp Sept 20


' 1. poprawiony on focus
' 2. poprawione przekopiowanie dat i cw miedzy formularzem a arkuszem details
' 3. dodanie update rowniez przy cofaniu
' 4. przesuniecie petli o jeden stopen dzieki czemu mozemy rowniez dodawac dane do ostatniej komorki details


' ===================================================


' teraz troche przerwy jesli chodzi wersjonowanie

' 42262_5844675926
' edytuj_projekt - new button



' 42262_5806481481
' wersja z dodatkowym enumem

' 42262_5772685185
' ===================================================
' zbior enum zostal przeniesiony do osobnego modulu
' z modulu Global w ktorym teraz z sensem tylko
' pojedyncze zmienne dostepne
' ===================================================

' 42262_5740277778 version
' ===================================================

' wersja ktora dotychczasowo korzysta tylko i wylacznie z enuma
' zatem wszystkie zapiski znajdujace sie w arkuszu register nie maja zadnej
' sily sprawczej i stoja tylko i wylacznie na lepsze czasy
' gdyby enum okazal sie niewystarczajacym narzedziem
' dla pracy na danych
' ===================================================



