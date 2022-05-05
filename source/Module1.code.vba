Sub Nacti_data()

'------------------------------------
' Natahování dat z ARESu a ESM by SV
'------------------------------------
'Verze 1.0 - 5.5.2022

'Pozor na denní limit dotazů na ARES
'Účelem provozování aplikace je poskytnout rychlé a obecně dostupné informace o jednotlivých subjektech.
'K tomuto účelu není pro běžného uživatele přístup k aplikaci ARES omezen.
'S ohledem na charakter provozu ARES a jeho zabezpečení
'si Ministerstvo financí vyhrazuje právo omezit nebo znemožnit přístup
'k www aplikaci ARES uživatelům, kteří:
' - odešlou k vyřízení více než 10.000 dotazů v době od 8:00 hod. do 18:00 hod.,
' -odešlou k vyřízení více než 50.000 dotazů v době od 18:00 hod. do 8:00 hod. rána následujícího dne,
' -se snaží o porušení bezpečnostní ochrany www serveru Ministerstva financí,
' -opakovaně posílají nesprávně vyplněné dotazy,
' -opakovaně posílají stejné dotazy,
' -mají větší počet současně zadaných dotazů (pro automatizované XML dotazy),
' -obcházejí povolené limity využíváním dotazování z většího množství IP adres,
' -automatizovaně propátrávají databázi náhodnými údaji nebo generují většinu nesprávných dotazů.

'Definice proměnných
Dim i, j As Long 'Počítadla
Dim Odkaz As String
Dim IC As Long
Dim Oblast As Range
Dim Majitel As Long
Dim PocetMajitelu As Long
Dim StringPocetMajitelu, KolikatyMajitel As String

'Načtení seznamu IČ z 1.listu a 1. sloupce - DATA
'Smyčka pro IČ z listu DATA
For i = 2 To Sheet1.Cells(Rows.Count, 1).End(xlUp).Row

  'Odkaz na web ARES
  Odkaz = "URL;http://wwwinfo.mfcr.cz/cgi-bin/ares/darv_or.cgi?ico=" & Sheet1.Cells(i, 1) & "&jazyk=cz&xml=2"

  'Dotaz a jeho uložení na Sheet2
  With Sheet2.Range("a1").QueryTable
     .Connection = Odkaz
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
  End With


  'Kontrola, zda je možné najít IČO
  Set Oblast = Sheet2.Cells.Find(What:="IČO:", LookIn:=xlFormulas, LookAt:= _
          xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
          , SearchFormat:=False)

  If Not Oblast Is Nothing Then

    'Přenesení hodnot
    IC = Sheet2.Range("a:a").Find(What:="IČO:", LookAt:=xlWhole).Row

   'Název subjektu
    Sheet1.Cells(i, 2) = Sheet2.Cells(IC + 1, 2)

   'Adresa
    Sheet1.Cells(i, 3) = Sheet2.Cells(IC + 3, 2)

  Else
    'Zapíše info, že subjekt nenalezl
    Sheet1.Cells(i, 2) = "NENALEZENO"
  End If


  'Odkaz na web ESM
  Odkaz = "URL;https://esm.justice.cz/ias/issm/rejstrik-$sm?p%3A%3Asubmit=x&.%2Frejstrik-%24sm=&nazev=&ico=" & Sheet1.Cells(i, 1) & "&oddil=&vlozka=&soud=&polozek=50&typHledani=STARTS_WITH"

  'Dotaz a jeho uložení na Sheet2 (WEB_down)
  With Sheet2.Range("a1").QueryTable
     .Connection = Odkaz
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
  End With

  'Kontrola, zda je možné najít nějaké zapsané skutečné majitele
  Set Oblast = Sheet2.Cells.Find(What:="1. Jméno:", LookIn:=xlFormulas, LookAt:= _
          xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
          , SearchFormat:=False)

  If Not Oblast Is Nothing Then

  'Přenesení hodnot
  'Nalezení řetězce Počet nalezených skutečných majitelů: kde pak následuje počet majitelů

    Majitel = Sheet2.Range("a:a").Find(What:="Počet nalezených skutečných majitelů:", LookAt:=xlPart).Row
    StringPocetMajitelu = Sheet2.Cells(Majitel, 1)

    'Vykuchání hodnoty počtu majitelů
    PocetMajitelu = Mid(StringPocetMajitelu, 39, 1)


    For j = 1 To PocetMajitelu
        'Složení řetězce pro hledání dle počtu majitelů
        KolikatyMajitel = j & ". Jméno:"
        Majitel = Sheet2.Range("a:a").Find(What:=KolikatyMajitel, LookAt:=xlPart).Row

        'Majitel číslo j a jeho jméno přepsané do seznamu
        Sheet1.Cells(i, j + 3) = Sheet2.Cells(Majitel, 2)
    Next j

  Else
    'Zapíše info, že k subjekt nenalezl skutečné majitele
    Sheet1.Cells(i, 4) = "NENALEZENO"
  End If


Next i

'Konec hledání
MsgBox ("Hledání dokončeno.")
End Sub
