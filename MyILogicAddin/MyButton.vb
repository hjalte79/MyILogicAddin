Imports Inventor
Public Class MyButton
    Private _inventor As Inventor.Application
    Private _settingsButton As ButtonDefinition
    Public Sub New(inventor As Inventor.Application)
        _inventor = inventor

        SetupButtonDefinition()
        AddButtonDefinitionToRibbon()
    End Sub

    Private Sub SetupButtonDefinition()

        Dim conDefs As ControlDefinitions = _inventor.CommandManager.ControlDefinitions
        _settingsButton = conDefs.AddButtonDefinition(
            "MyButton DisplayName",
            "MyButton InternalName",
            CommandTypesEnum.kEditMaskCmdType,
            Guid.NewGuid().ToString(),
            "MyButton DescriptionText",
            "MyButton ToolTipText",)

        AddHandler _settingsButton.OnExecute, AddressOf MyButton_OnExecute

        _settingsButton.StandardIcon = PictureDispConverter.ToIPictureDisp(My.Resources.MyImage16x16)
        _settingsButton.LargeIcon = PictureDispConverter.ToIPictureDisp(My.Resources.MyImage32x32)
    End Sub

    Private Sub AddButtonDefinitionToRibbon()

        Dim ribbon As Ribbon = _inventor.UserInterfaceManager.Ribbons.Item("Assembly")
        Dim ribbonTab As RibbonTab = ribbon.RibbonTabs.Item("id_TabManage")
        Dim ribbonPanel As RibbonPanel = ribbonTab.RibbonPanels.Item("iLogic.RibbonPanel")
        ribbonPanel.CommandControls.AddButton(_settingsButton)

    End Sub

    Private Sub MyButton_OnExecute(Context As NameValueMap)

        Try

            Dim rule As New ThisRule()
            rule.ThisApplication = _inventor
            rule.Main()

        Catch ex As Exception

            MsgBox("Something went wrong while runing rule. Message: " & ex.Message)

        End Try

    End Sub

End Class
