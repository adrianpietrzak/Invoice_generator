Sub WydrukDoPDF()

Dim wsA As Worksheet
Dim wbA As Workbook
Dim strTime As String
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile As Variant
On Error GoTo errHandler

Set wbA = ActiveWorkbook
Set wsA = ActiveSheet
strTime = Format(Now(), "yyyymmdd\_hhmm")

'pobiera sciezke do aktywnego workbooka
strPath = wbA.Path
If strPath = "" Then
  strPath = Application.DefaultFilePath
End If
strPath = strPath & "\"

'zamienia niedozwolone znaki w nazwie pliku
strName = Replace(wsA.Name, " ", "")
strName = Replace(strName, ".", "_")
rodzajFaktury = Range("J6").Value
numerFaktury = Replace(Range("J1").Value, "/", "_")

'tworzenie nazwy pliku do zapisania
strFile = "Faktura" & "_" & rodzajFaktury & "_" & numerFaktury & ".pdf"
strPathFile = strPath & strFile



'mozesz wpisac nazwe i wybrac sciezke do pliku
myFile = Application.GetSaveAsFilename _
    (InitialFileName:=strPathFile, _
        Filefilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Folder and FileName to save")

'eksport do PDF jezeli folder zostal wybrany
If myFile <> "False" Then
    wsA.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=myFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
'wiadomosc potwierdzajaca
    MsgBox "Faktura została zapisana do pliku PDF: " _
      & vbCrLf _
      & myFile
End If

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Nie mogę stworzyć pliku PDF"
    Resume exitHandler
End Sub

