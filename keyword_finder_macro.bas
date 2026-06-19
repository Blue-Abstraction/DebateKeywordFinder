' CHANGE THIS TO THE FULL PATH OF YOUR KEYWORD TEXT FILE
Const KEYWORD_FILE As String = "YOUR\PATH\HERE\debate_words.txt"

Public Sub ScanDocument()

    Dim keywords As Collection
    Dim keyword As Variant

    Dim fileNum As Integer
    Dim lineText As String

    Dim searchRange As Range
    Dim totalMatches As Long

    Dim originalDoc As Document
    Dim doc As Document

    Dim newPath As String
    Dim baseFileName As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set originalDoc = ActiveDocument

    If originalDoc.Path = "" Then
        MsgBox "Please save the document first.", vbExclamation
        GoTo CleanUp
    End If

    originalDoc.Save

    baseFileName = "[K] " & Left(originalDoc.Name, InStrRev(originalDoc.Name, ".") - 1) & ".docx"
    newPath = GetDownloadsPath() & baseFileName

    Call CloseDocumentIfOpen(baseFileName)

    Set doc = Documents.Add(originalDoc.FullName)
    doc.SaveAs2 FileName:=newPath, FileFormat:=wdFormatDocumentDefault

    Set keywords = New Collection

    fileNum = FreeFile
    Open KEYWORD_FILE For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)

        If lineText <> "" Then
            keywords.Add lineText
        End If
    Loop

    Close #fileNum

    For Each keyword In keywords

        Set searchRange = doc.Content

        With searchRange.Find
            .ClearFormatting
            .Text = CStr(keyword)

            .MatchCase = False
            .MatchWholeWord = True

            .Forward = True
            .Wrap = wdFindStop

            Do While .Execute

                totalMatches = totalMatches + 1

                doc.Comments.Add _
                    Range:=searchRange.Duplicate, _
                    Text:="Keyword found: " & CStr(keyword)

                searchRange.Collapse wdCollapseEnd

            Loop
        End With

    Next keyword

    doc.Save

    MsgBox totalMatches & " occurrences found." & vbCrLf & _
           "Saved as:" & vbCrLf & newPath

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Public Function GetDownloadsPath() As String
    Dim downloadsPath As String
    Dim username As String

#If Mac Then
    username = Environ("USER")
    downloadsPath = "/Users/" & username & "/Downloads/"
#Else
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    downloadsPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
#End If

    GetDownloadsPath = downloadsPath
End Function

Public Sub CloseDocumentIfOpen(docName As String)
    On Error Resume Next
    Documents(docName).Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub

