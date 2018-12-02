Sub generator()


'otwieranie bazy klientow
Dim bazaklientow As Excel.Workbook
Set bazaklientow = Workbooks.Open("C:\Users\Michał\Desktop\Program Faktura\Baza_klientow.xlsx")


'wyszukiwanie po nazwie firmy
firma = InputBox("Podaj nazwe wyszukiwanej osoby bądź firmy", "Nazwa Osoby/Firmy")

If Not firma = "" Then
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4, Criteria1:="=*" & firma & "*"


'wyszukiwane po nipie
    Else
    NIP = InputBox("Podaj NIP wyszkuwanej osoby bądź firmy. Aby makro działało poprawnie podaj PEŁNY NIP", "NIP")

            If firma = "" And NIP = "" Then
            MsgBox ("Brak wymaganych danych. Wystaw fakturę ponownie")
            bazaklientow.Close False
            Exit Sub
            End If
            
            ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=2, Criteria1:=NIP

End If


'tworzenie nowej faktury
Set Nowafaktura = Workbooks.Add
ActiveSheet.Name = "Faktura"
Workbooks("Faktura_template.xlsm").Worksheets("Faktura_template").Activate
Range("A1", "K49").Copy
Nowafaktura.Activate
Range("A1").PasteSpecial
Range("A1").PasteSpecial Paste:=xlPasteColumnWidths

'ustawienie marginesow
 Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.236220472440945)
        .RightMargin = Application.InchesToPoints(0.236220472440945)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    

Workbooks("Faktura_template.xlsm").Worksheets("FormyPlatnosci").Activate
Range("A1", "C3").Copy


'dodawanie listy rozwijalnej
Nowafaktura.Activate
Worksheets.Add().Name = "FormyPlatnosci"
Range("A1").PasteSpecial Paste:=xlPasteValues
Worksheets("Faktura").Activate
Range("D37:F37").Select
With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=FormyPlatnosci!$A$2:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
End With

bazaklientow.Activate
Selection.SpecialCells(xlCellTypeVisible).Copy

Nowafaktura.Activate
Worksheets.Add().Name = "baza"
Range("A1").PasteSpecial Paste:=xlPasteValues



'uzupelnianie faktury danymi

'nip
Range("B2").Copy
Worksheets("Faktura").Activate
Range("J14:K14").Select
ActiveSheet.Paste

'imie
Worksheets("baza").Activate
Range("D2").Copy
Worksheets("Faktura").Activate
Range("I11:K11").Select
ActiveSheet.Paste

'ulica
Worksheets("baza").Activate
Range("E2").Copy
Worksheets("Faktura").Activate
Range("I12:K12").Select
ActiveSheet.Paste

'kod pocztowy i miasto
Worksheets("baza").Activate
Range("J2") = " "
Range("I2").Value = "=F2&J2&G2"
Range("I2").Copy
Range("I2").PasteSpecial xlPasteValues
Worksheets("Faktura").Activate
Range("I13:K13").Select
ActiveSheet.Paste

'logo i strona www
Workbooks("Faktura_template.xlsm").Worksheets("Faktura_template").Activate
Range("A1:F5").Copy
Nowafaktura.Activate
Worksheets("Faktura").Activate
Range("A1:F5").Select
ActiveSheet.Paste


'wstawianie sumy brutto
Workbooks("Faktura_template.xlsm").Worksheets("Faktura_template").Activate
Range("D34:E34").Copy
Nowafaktura.Activate
Worksheets("Faktura").Activate
Range("D34").PasteSpecial Paste:=xlPasteFormulas



'otwieranie pliku z historia faktur
Dim wystawionefaktury As Excel.Workbook
Set wystawionefaktury = Workbooks.Open("C:\Users\Michał\Desktop\Program Faktura\Wystawione_faktury.xlsx")

'wstawianie daty jako numeru faktury
Nowafaktura.Activate
Worksheets("Faktura").Activate

Dim Dzis As String
Dzis = Format(Now(), "ddmmyyyy")

Range("J1:K1") = Dzis

'pobor numeru faktury
wystawionefaktury.Activate
Worksheets("Faktury").Activate

OstKomorka = Cells(Rows.Count, "A").End(xlUp).Row
Str (OstKomorka)

ObecnyRok = Right(Dzis, 4)
OstatniaData = Range("A" & OstKomorka)
OstatniRok = Right(OstatniaData, 4)


If (ObecnyRok = OstatniRok) Then
               
'konwersja na wlasciwy numer faktury
Nowafaktura.Activate
Worksheets("Faktura").Activate
LiczbyDoZamiany = Left(Dzis, 2)
Range("J1:K1") = Replace(Dzis, LiczbyDoZamiany, OstKomorka)
Range("J1:K1").Select

NrFaktury = ActiveCell.Value

LiczbyZPrawejDoZamiany = Right(Dzis, 6)
ZmianaFormatuFaktury = "/" + Right(Dzis, 4)
Range("J1:K1") = Replace(NrFaktury, LiczbyZPrawejDoZamiany, ZmianaFormatuFaktury)


'tutaj zamienia date faktury na nowy rok
Else
Nowafaktura.Activate
Worksheets("Faktura").Activate
Range("J1:K1").Value = "1/" & ObecnyRok
End If


'zamyka bez zapisywania zmian
bazaklientow.Close False
wystawionefaktury.Close False


End Sub

'-------------------------------------------------------------------------------------------------------------------------------

Sub zatwierdz()
Dim Nowafaktura As Workbook
Set Nowafaktura = ActiveWorkbook


Worksheets("Faktura").Activate


'kasuje niepotrzebne spreadsheety
Application.DisplayAlerts = False
Worksheets("FormyPlatnosci").Delete
Worksheets("baza").Delete
Application.DisplayAlerts = True


numerFaktury = Replace(Range("J1").Value, "/", "_")

ActiveSheet.Name = "Faktura" & " " & numerFaktury

Range("J1:K1").Copy

'otwarcie pliku historycznego
Dim wystawionefaktury As Excel.Workbook
Set wystawionefaktury = Workbooks.Open("C:\Users\Michał\Desktop\Program Faktura\Wystawione_faktury.xlsx")

                        
                        
'uzupelnianie pliku z historia
                        
'kopiowanie nr faktury
OstKomorka = Cells(Rows.Count, "A").End(xlUp).Row
ObecnaFaktura = Range("A" & OstKomorka).Offset(1, 0).Select
ActiveSheet.Paste
Selection.UnMerge

'kopiowanie odbiorcy faktury
Nowafaktura.Activate

Range("I11:K11").Copy

wystawionefaktury.Activate
OstKomorka = Cells(Rows.Count, "A").End(xlUp).Row
ObecnaFaktura = Range("A" & OstKomorka).Offset(0, 1).Select
ActiveSheet.Paste
Selection.UnMerge

'kopiowanie kwoty faktury
Nowafaktura.Activate
Cells.Find(What:="Razem", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Offset(0, 1).Select
        
ActiveCell.Copy
        
wystawionefaktury.Activate
OstKomorka = Cells(Rows.Count, "A").End(xlUp).Row
Range("A" & OstKomorka).Offset(0, 2).PasteSpecial Paste:=xlPasteValues

Range("A" & OstKomorka).Offset(-1, 0).Copy
Range("A" & OstKomorka).PasteSpecial Paste:=xlPasteFormats

Range("B" & OstKomorka).Offset(-1, 0).Copy
Range("B" & OstKomorka).PasteSpecial Paste:=xlPasteFormats

Range("C" & OstKomorka).Offset(-1, 0).Copy
Range("C" & OstKomorka).PasteSpecial Paste:=xlPasteFormats

'kod zapisujacy fakture oryginał
Nowafaktura.Activate
Application.DisplayAlerts = False
            
Range("J1:K1").Copy
Range("F1").Select
ActiveSheet.Paste
NrZapisu = ActiveCell.Value
LiczbyZPrawejDoZamiany = Right(NrZapisu, 5)
ZmianaFormatuFaktury = "_" + Right(NrZapisu, 4)
Range("F1:G1") = Replace(NrZapisu, LiczbyZPrawejDoZamiany, ZmianaFormatuFaktury)
NrZapisu = ActiveCell.Value
Dim sciezka As String
sciezka = "C:\Users\Michał\Desktop\Program Faktura\Wystawione Faktury\"
nazwa_pliku = "Faktura_oryginał_" & NrZapisu & ".xlsx"
ActiveWorkbook.SaveAs Filename:=sciezka & nazwa_pliku
Range("F1:G1").Clear
ActiveWorkbook.Save


'wstawianie hiperłącza
Adres = sciezka & nazwa_pliku
wystawionefaktury.Activate
OstKomorka = Cells(Rows.Count, "D").End(xlUp).Row + 1
KomorkaAdresowa = Range("D" & OstKomorka)

ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=Range("D" & OstKomorka), _
                                   Address:=Adres, _
                                   TextToDisplay:="Oryginał"




'kod zapisujacy kopie faktury
Nowafaktura.Activate
Range("J6").Value = "Kopia"
Range("J1:K1").Copy
Range("F1").Select
ActiveSheet.Paste
NrZapisu = ActiveCell.Value
LiczbyZPrawejDoZamiany = Right(NrZapisu, 5)
ZmianaFormatuFaktury = "_" + Right(NrZapisu, 4)
Range("F1:G1") = Replace(NrZapisu, LiczbyZPrawejDoZamiany, ZmianaFormatuFaktury)
NrZapisu = ActiveCell.Value
nazwa_pliku = "Faktura_kopia_" & NrZapisu & ".xlsx"
ActiveWorkbook.SaveAs Filename:=sciezka & nazwa_pliku
Range("F1:G1").Clear
ActiveWorkbook.Save
            
Application.DisplayAlerts = True


Adres = sciezka & nazwa_pliku
wystawionefaktury.Activate
OstKomorka = Cells(Rows.Count, "E").End(xlUp).Row + 1
KomorkaAdresowa = Range("E" & OstKomorka)

ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=Range("E" & OstKomorka), _
                                   Address:=Adres, _
                                   TextToDisplay:="Kopia"
                                   


'zamykanie pliku z fakturami
wystawionefaktury.Close True
Nowafaktura.Close False

End Sub
