Sub fly_away()
    ' Store indiviualized ready-to-send mails with different attachments as
    ' drafts in Outlook with name and mail address given  in Excel spreadsheet.
    ' Enter path to attachements here.
    '
    ' Must add references in Excel to
    ' 1) OLE Automation
    ' 2) Microsoft Outlook xx.0 Object Library
    Dim OutApp As Object
    Dim OutMail As Object
    Dim OutAccountExplorers As Object
    Dim cell As Range
    Dim strPath As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutAccountExplorers = OutApp.Session.Accounts.Item(2) ' Use secondary account

    strPath = "C:\Users\path\to\pdfs"

    ' No error handling
    For Each cell In Columns("A").Cells.SpecialCells(xlCellTypeConstants)
         i = cell.Row
         Set OutMail = OutApp.CreateItem(0)
         With OutMail
             .To = Sheets("Personenliste").Range("C" & i).Value
             .Subject = "Enter Subject Here"
             .Body = "Hello " & Sheets("Personenliste").Range("A" & i).Value & "," & vbNewLine & vbNewLine & "welcome to planet future!"
             .Attachments.Add strPath & Sheets("Personenliste").Range("B" & i).Value & "_" & Sheets("Personenliste").Range("A" & i).Value & ".pdf"
            Set .SendUsingAccount = OutAccountExplorers
             .Close (olSave)
             '.Display
         End With
         Set OutMail = Nothing
    Next cell
End Sub
