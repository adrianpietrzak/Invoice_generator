Sub Proforma()
'otwieranie pliku z faktura
Faktura = Application.GetOpenFilename(Title:="Otwórz właściwą fakturę")
If Faktura <> False Then
        Workbooks.Open (Faktura)
    Else
        Exit Sub
End If

'zmiana nazwy pliku na PROFORMA
Range("J6") = "Proforma"


'zapis faktury
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
nazwa_pliku = "Faktura_ProForma_" & NrZapisu & ".xlsx"
ActiveWorkbook.SaveAs Filename:=sciezka & nazwa_pliku
Range("F1:G1").Clear
ActiveWorkbook.Save
Application.DisplayAlerts = True

End Sub
