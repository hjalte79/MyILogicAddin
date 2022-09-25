Imports System.Runtime.InteropServices
Imports Inventor

' Change the GUID here and use the same as in the addin fiel!
<GuidAttribute("a7558ed7-8d2d-41bb-9295-4f3c9fe5d902"), ComVisible(True)>
Public Class StandardAddInServer
    Implements Inventor.ApplicationAddInServer

    Private _myButton As MyButton

    ''' <summary>
    '''     Invoked by Autodesk Inventor after creating the AddIn. 
    '''     AddIn should initialize within this call.
    ''' </summary>
    ''' <param name="AddInSiteObject">
    '''     Input argument that specifies the object, which provides 
    '''     access to the Autodesk Inventor Application object.
    ''' </param>
    ''' <param name="FirstTime">
    '''     The FirstTime flag, if True, indicates to the Addin that this is the 
    '''     first time it is being loaded and to take some specific action.
    ''' </param>
    Public Sub Activate(AddInSiteObject As ApplicationAddInSite, FirstTime As Boolean) Implements ApplicationAddInServer.Activate
        Try

            ' initialize the rule class
            _myButton = New MyButton(AddInSiteObject.Application)

        Catch ex As Exception

            ' Show a message if any thing goes wrong.
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    ''' <summary>
    '''     Invoked by Autodesk Inventor to shut down the AddIn. 
    '''     AddIn should complete shutdown within this call.
    ''' </summary>
    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        _myButton.Unload()
        _myButton = Nothing
    End Sub

    ''' <summary>
    '''     Invoked by Autodesk Inventor in response to user requesting the execution 
    '''     of an AddIn-supplied command. AddIn must perform the command within this call.
    ''' </summary>
    Public Sub ExecuteCommand(CommandID As Integer) Implements ApplicationAddInServer.ExecuteCommand

    End Sub

    ''' <summary>
    '''     Gets the IUnknown of the object implemented inside the AddIn that supports AddIn-specific API.
    ''' </summary>
    Public ReadOnly Property Automation As Object Implements ApplicationAddInServer.Automation
        Get
            Throw New NotImplementedException()
        End Get
    End Property
End Class
