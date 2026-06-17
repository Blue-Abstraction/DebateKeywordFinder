Sub ScanDocument()

    Dim keywords As Collection
    Dim keyword As Variant

    Dim fileNum As Integer
    Dim lineText As String

    Dim searchRange As Range
    Dim totalMatches As Long

    Dim originalDoc As Document
    Dim doc As Document

    Dim newPath As String
    Dim fso As Object

    ' ===== KEEP ORIGINAL SAFE =====

    Set originalDoc = ActiveDocument

    If originalDoc.Path = "" Then
        MsgBox "Please save the document first."
        Exit Sub
    End If

    newPath = Environ("USERPROFILE") & _
        "\Downloads\" & _
        "[K] " & Left(originalDoc.Name, InStrRev(originalDoc.Name, ".") - 1) & _
        ".docx"

    originalDoc.Save

    ' ===== SAFE FILE COPY (NO WORD LOCK ISSUES) =====

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(newPath) Then
        On Error Resume Next
        Kill newPath
        On Error GoTo 0
    End If

    fso.CopyFile originalDoc.FullName, newPath, True

    ' ===== OPEN COPY ONLY =====

    Set doc = Documents.Open(newPath)

    ' ===== LOAD KEYWORDS =====

    Set keywords = New Collection

    fileNum = FreeFile

    Open "INSERT\YOUR\FILE\NAME\HERE.txt" For Input As #fileNum

    Do While Not EOF(fileNum)

        Line Input #fileNum, lineText

        lineText = Trim(lineText)

        If lineText <> "" Then
            keywords.Add lineText
        End If

    Loop

    Close #fileNum

    ' ===== SEARCH COPY =====

    Application.ScreenUpdating = False

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

    Application.ScreenUpdating = True

    doc.Save

    MsgBox totalMatches & " occurrences found." & vbCrLf & _
           "Saved as:" & vbCrLf & newPath

End Sub

