Imports System.Text

Public Class WindowsEnumerator
    Private Delegate Function EnumCallBackDelegate(hwnd As Integer, lParam As Integer) As Integer
    Private Declare Function EnumWindows Lib "user32" (lpEnumFunc As EnumCallBackDelegate, lParam As Integer) As Integer
    Private Declare Function EnumChildWindows Lib "user32" (hWndParent As Integer, lpEnumFunc As EnumCallBackDelegate, lParam As Integer) As Integer
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (hwnd As Integer, lpClassName As StringBuilder, nMaxCount As Integer) As Integer
    Private Declare Function IsWindowVisible Lib "user32" (hwnd As Integer) As Integer
    Private Declare Function GetParent Lib "user32" (hwnd As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As Int32, wMsg As Int32, wParam As Int32, lParam As Int32) As Int32
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As Int32, wMsg As Int32, wParam As Int32, lParam As StringBuilder) As Int32
    Private _listChildren As New List(Of ApiWindow)
    Private _listTopLevel As New List(Of ApiWindow)
    Private _topLevelClass As String = ""
    Private _childClass As String = ""

    Public Overloads Function GetTopLevelWindows() As List(Of ApiWindow)
        EnumWindows(AddressOf EnumWindowProc, &H0)
        Return _listTopLevel
    End Function

    Public Overloads Function GetTopLevelWindows(className As String) As List(Of ApiWindow)
        _topLevelClass = className
        Return Me.GetTopLevelWindows()
    End Function

    Public Overloads Function GetChildWindows(hwnd As Int32) As List(Of ApiWindow)
        _listChildren = New List(Of ApiWindow)
        EnumChildWindows(hwnd, AddressOf EnumChildWindowProc, &H0)
        Return _listChildren
    End Function

    Public Overloads Function GetChildWindows(hwnd As Int32, childClass As String) As List(Of ApiWindow)
        _childClass = childClass
        Return Me.GetChildWindows(hwnd)
    End Function
    Private Function EnumWindowProc(hwnd As Int32, lParam As Int32) As Int32
        If GetParent(hwnd) = 0 AndAlso CBool(IsWindowVisible(hwnd)) Then
            Dim window As ApiWindow = GetWindowIdentification(hwnd)
            If _topLevelClass.Length = 0 OrElse window.ClassName.ToLower() = _topLevelClass.ToLower() Then
                _listTopLevel.Add(window)
            End If
        End If
        Return 1
    End Function

    Private Function EnumChildWindowProc(hwnd As Int32, lParam As Int32) As Int32
        Dim window As ApiWindow = GetWindowIdentification(hwnd)
        If _childClass.Length = 0 OrElse window.ClassName.ToLower() = _childClass.ToLower() Then
            _listChildren.Add(window)
        End If
        Return 1
    End Function

    Private Function GetWindowIdentification(hwnd As Integer) As ApiWindow
        Const WM_GETTEXT = &HD
        Const WM_GETTEXTLENGTH = &HE
        Dim window As New ApiWindow()
        Dim title As New StringBuilder()
        Dim size As Int32 = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
        If size > 0 Then
            title = New StringBuilder(size + 1)
            SendMessage(hwnd, WM_GETTEXT, title.Capacity, title)
        End If
        Dim classBuilder As New StringBuilder(64)
        GetClassName(hwnd, classBuilder, 64)
        window.ClassName = classBuilder.ToString()
        window.MainWindowTitle = title.ToString()
        window.hWnd = hwnd
        Return window
    End Function

    Public Class ApiWindow
        Public MainWindowTitle As String = ""
        Public ClassName As String = ""
        Public hWnd As Int32
    End Class

End Class
