Imports Inventor

Public Module DockableWindowChildAdapter
    ''' <summary>
    ''' https://forums.autodesk.com/t5/inventor-forum/dockable-window-with-wpf-controls-don-t-receive-keyboard-input/td-p/9115997
    ''' </summary>
    Private Const DLGC_WANTARROWS As UInt32 = &H1
    Private Const DLGC_WANTTAB As UInt32 = &H2
    Private Const DLGC_WANTALLKEYS As UInt32 = &H4
    Private Const DLGC_HASSETSEL As UInt32 = &H8
    Private Const DLGC_WANTCHARS As UInt32 = &H80
    Private Const WM_GETDLGCODE As UInt32 = &H87

    Private _dockableWindow As DockableWindow

    Public Sub AddWPFWindow(dockableWindow As DockableWindow, window As Window)
        _dockableWindow = dockableWindow

        window.WindowStyle = WindowStyle.None
        window.WindowState = WindowState.Maximized
        window.ResizeMode = ResizeMode.NoResize
        window.Show()

        Dim win As Window = Window.GetWindow(window)
        Dim wih = New Interop.WindowInteropHelper(win)
        Dim hWnd As IntPtr = wih.EnsureHandle()

        _dockableWindow.AddChild(hWnd)

        Interop.HwndSource.FromHwnd(hWnd).AddHook(AddressOf WndProc)
    End Sub

    Public Function WndProc(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        If (msg = WM_GETDLGCODE) Then

            handled = True
            Return New IntPtr(DLGC_WANTCHARS Or DLGC_WANTARROWS Or DLGC_HASSETSEL Or DLGC_WANTTAB Or DLGC_WANTALLKEYS)
        End If
        Return IntPtr.Zero
    End Function
End Module