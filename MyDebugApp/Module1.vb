Imports Inventor
Imports System.Runtime.InteropServices
Imports MyILogicAddin ' This is the namespace from the addin dll!

Module Module1

    Sub Main()
        Dim inventorObject As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        Dim rule As New ThisRule()
        rule.ThisApplication = inventorObject
        rule.Main()
    End Sub

End Module
