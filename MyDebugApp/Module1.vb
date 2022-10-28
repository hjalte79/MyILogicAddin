Imports Inventor
Imports System.Runtime.InteropServices
Imports MyILogicAddin ' This is the namespace from the addin dll!

Module Module1

    Sub Main()
        Dim inventorObject As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        Dim mySearchForm As New MySearchForm(inventorObject)
        mySearchForm.ShowDialog()

    End Sub

End Module
