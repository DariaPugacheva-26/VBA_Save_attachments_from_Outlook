Attribute VB_Name = "Save_attachments"
Option Explicit

Sub Save_attachments_from_Outlook()

Dim Ex As New Excel.Application 
'i needed to choose folder where i want to save files.
'in Outlook i didn't find such option, but did find in Excel
'so i open Excel to choose folder.
 
Dim Path As String
Dim Path2 As String
Dim SLCT As Outlook.Selection
Dim Mails As Outlook.MailItem
Dim Att As Outlook.Attachments
Dim i As Long
Dim AttCount As Long

With Ex.Application.FileDialog(msoFileDialogFolderPicker)
'open excel Folder picker Dialog and selecting folder

    If .Show = -1 Then
        Path = .SelectedItems(1)
    Else: GoTo Handle 'if canceled
    End If
End With

On Error GoTo Handle2
If Path = "" Then GoTo Handle

Set SLCT = Application.ActiveExplorer.Selection
For Each Mails In SLCT
    Set Att = Mails.Attachments
    AttCount = Att.Count
        If AttCount > 0 Then
            For i = AttCount To 1 Step -1
                Path2 = Path & "\" & Att.Item(i).FileName
                Att.Item(i).SaveAsFile Path2
            Next i
        End If
Mails.UnRead = False
Next Mails
Handle:

Ex.Quit
Exit Sub

Handle2:
MsgBox "Something went wrong, try again!"
Ex.Quit
End Sub

