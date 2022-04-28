Imports Inventor

Public Class ThisRule
    Inherits AbstractRule

    Private searchText As String
    Private iLogicAddinGuid As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
    Private iLogicAddin As ApplicationAddIn = Nothing
    Private iLogicAutomation = Nothing
    Private outputFile As String = "c:\TEMP\seachedRules.txt"

    Sub Main()
        If (IO.File.Exists(outputFile)) Then
            IO.File.Delete(outputFile)
        End If
        searchText = InputBox("Text to search for", "Search")

        iLogicAddin = ThisApplication.ApplicationAddIns.ItemById(
            "{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
        iLogicAutomation = iLogicAddin.Automation
        Dim doc As AssemblyDocument = ThisDoc.Document

        searchDoc(doc)
        For Each refDoc As Document In doc.AllReferencedDocuments
            searchDoc(refDoc)
        Next

        Process.Start("notepad.exe", outputFile)
    End Sub

    Private Sub searchDoc(doc As Document)
        Dim rules = iLogicAutomation.Rules(doc)
        If (rules Is Nothing) Then Return
        For Each rule In rules
            Dim strReader As IO.StringReader = New IO.StringReader(rule.Text)
            Dim i As Integer = 1

            Do While (True)
                Dim line As String
                line = strReader.ReadLine()
                If line Is Nothing Then Exit Do
                If (line.ToUpper().Contains(searchText.ToUpper())) Then
                    Dim nl = System.Environment.NewLine
                    IO.File.AppendAllText(outputFile,
                        "Doc name : " & doc.DisplayName & nl &
                        "Rule name: " & rule.Name & nl &
                        "line " & i & "  : " & line.Trim() & nl & nl)
                End If
                i = i + 1
            Loop
        Next
    End Sub
End Class
