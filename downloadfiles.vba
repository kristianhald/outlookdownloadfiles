Public Sub Download()
    Dim item As Object
    Dim oAttachment As Outlook.Attachment
    Dim sSaveFolder As String
    Dim iCount As Integer
    Dim attExt As String

    iCount = 0
    sSaveFolder = GetDesktop & "\Faktura-" & Year(Now()) & Month(Now()) & Day(Now()) & "-" & Hour(Now()) & Minute(Now())
    CreateFolderIfMissing sSaveFolder
    For Each item In Application.ActiveExplorer.Selection
        For Each oAttachment In item.Attachments
            attExt = UCase(Right(oAttachment.FileName, 4))
            If attExt = ".DOC" Or attExt = "DOCX" Or attExt = ".ODT" Or attExt = ".RTF" Or attExt = ".TXT" Or attExt = ".WPD" Or attExt = ".WPS" Or attExt = ".CSV" Or attExt = ".PPS" Or attExt = ".PPT" Or attExt = "PPTX" Or attExt = ".PDF" Or attExt = ".XLR" Or attExt = ".XLS" Or attExt = "XLSX" Or attExt = ".HTM" Or attExt = "HTML" Then
                oAttachment.SaveAsFile sSaveFolder & "\" & oAttachment.DisplayName
                iCount = iCount + 1
            End If
        Next
    Next
    
    MsgBox "Alle valgte vedhæftet filer hentet. " & iCount & " hentet."
End Sub

Function GetDesktop() As String
    Dim oWSHShell As Object

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    Set oWSHShell = Nothing
End Function

Sub CreateFolderIfMissing(path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then
       MkDir path
    End If
End Sub
