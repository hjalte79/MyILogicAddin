Imports System.Text
Imports Inventor

Public Class MySearchForm

    Private ReadOnly iLogicAddinGuid As String = "{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}"
    Private ReadOnly _iLogicAutomation
    Private ReadOnly _nl As String
    Private ReadOnly _inventor As Inventor.Application

    Public Sub New(inventor As Inventor.Application)

        ' This call is required by the designer.
        InitializeComponent()

        _inventor = inventor
        Dim iLogicAddin As ApplicationAddIn = _inventor.ApplicationAddIns.ItemById(iLogicAddinGuid)
        _iLogicAutomation = iLogicAddin.Automation

        _nl = System.Environment.NewLine

    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim _document As Document = _inventor.ActiveDocument
        Dim stringBuilder As New StringBuilder()

        SearchDoc(_Document, stringBuilder)
        For Each refDoc As Document In _Document.AllReferencedDocuments
            SearchDoc(refDoc, stringBuilder)
        Next

        TbResults.Text = stringBuilder.ToString()
    End Sub


    Private Sub SearchDoc(doc As Document, stringBuilder As StringBuilder)
        Dim searchText = TbSearchText.Text
        Dim rules = _iLogicAutomation.Rules(doc)
        If (rules Is Nothing) Then Return
        For Each rule In rules
            Dim strReader As New IO.StringReader(rule.Text)
            Dim i As Integer = 1

            Do While (True)
                Dim line As String
                line = strReader.ReadLine()
                If line Is Nothing Then Exit Do
                If (line.ToUpper().Contains(searchText.ToUpper())) Then

                    stringBuilder.Append($"Doc name : {doc.DisplayName}{_nl}")
                    stringBuilder.Append($"Rule name: {rule.Name}{_nl}")
                    stringBuilder.Append($"Line {i}: {line.Trim()}{_nl}")
                    stringBuilder.Append(_nl)

                End If
                i += 1
            Loop
        Next
    End Sub

End Class
