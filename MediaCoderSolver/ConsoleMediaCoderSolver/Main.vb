Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.JScript
Imports Microsoft.JScript.Vsa

Module Main
    #Region "WINAPI & Helpers"
    <DllImport("USER32.DLL")> Private Function GetShellWindow() As IntPtr
    End Function

    <DllImport("USER32.DLL")> Private Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("USER32.DLL")> Private Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)> Private Function GetWindowThreadProcessId(hWnd As IntPtr, <Out> ByRef lpdwProcessId As UInt32) As UInt32
    End Function

    <DllImport("USER32.DLL")> Private Function IsWindowVisible(hWnd As IntPtr) As Boolean
    End Function

    Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As Integer) As Boolean

    <DllImport("USER32.DLL")> Private Function EnumWindows(enumFunc As EnumWindowsProc, lParam As Integer) As Boolean
    End Function

    Private ReadOnly HShellWindow As IntPtr = GetShellWindow()
    Private ReadOnly DictWindows As New Dictionary(Of IntPtr, String)
    Private _currentProcessId As Integer
    Const WM_SETTEXT = &HC
    Const WM_GETTEXT = &HD
    Const WM_GETTEXTLENGTH = &HE
    Const BM_CLICK = &HF5

    Declare Auto Function SendMessage Lib "user32.dll" (hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    <DllImport("user32.dll")>
    Public Function FindWindowEx(parentHandle As IntPtr, childAfter As IntPtr, lclassName As String, windowTitle As String) As IntPtr
    End Function

     Public Declare Function SendMessageSTRING Lib "user32.dll" Alias "SendMessageA" (hwnd As Int32, wMsg As Int32, wParam As Int32, lParam As String) As Int32

        Public Function GetOpenWindowsFromPid(processId As Integer) As IDictionary(Of IntPtr, String)
        dictWindows.Clear()
        _currentProcessId = processID
        EnumWindows(AddressOf enumWindowsInternal, 0)
        Return dictWindows
    End Function

    Private Function EnumWindowsInternal(hWnd As IntPtr, lParam As Integer) As Boolean
        If (hWnd <> hShellWindow) Then
            Dim windowPid As UInt32
            If Not IsWindowVisible(hWnd) Then
                Return True
            End If
            Dim length As Integer = GetWindowTextLength(hWnd)
            If (length = 0) Then
                Return True
            End If
            GetWindowThreadProcessId(hWnd, windowPid)
            If (windowPid <> _currentProcessId) Then
                Return True
            End If
            Dim stringBuilder As New StringBuilder(length)
            GetWindowText(hWnd, stringBuilder, (length + 1))
            dictWindows.Add(hWnd, stringBuilder.ToString)
        End If
        Return True
    End Function

    #End Region

    Sub Main()
        On Error Resume Next
        Dim windowHandle As IntPtr = GetWantedWindowHandle()
        Dim labelText As String = GetLabelText(windowHandle)
        Dim answer As String = GetSumAnswer(labelText)
        SetTextboxText(windowHandle, answer)
        ClickButton(windowHandle)
    End Sub
    Private Function GetWantedWindowHandle() As IntPtr
        Dim wHandle As IntPtr
        For Each p As Process In Process.GetProcesses
            Dim windows As IDictionary(Of IntPtr, String) = GetOpenWindowsFromPID(p.Id)
            For Each kvp As KeyValuePair(Of IntPtr, String) In From kvp1 In windows Where kvp1.Value.ToLower = "start" = False Where kvp1.Value.StartsWith("Continue in ")
                wHandle = FindWindowEx(0, 0, Nothing, kvp.Value)
                Return wHandle
            Next
        Next
        Return Nothing
    End Function
    Private Function GetLabelText(wantedWindowHandle As IntPtr) As String
        Return (From itm In From itm1 In New WindowsEnumerator().GetChildWindows(WantedWindowHandle) Where itm1.MainWindowTitle.Contains("a simple math") Select itm1.MainWindowTitle).FirstOrDefault()
    End Function

    Private Function GetSumAnswer(workingText As String) As String
        Return CalculateFormula(workingText.Replace("Let's do a simple math: ", "").Replace(" ", "").Replace("=?", ""))
    End Function

    Public Function CalculateFormula(evaluationString as String) As String
        return Eval.JScriptEvaluate(evaluationString, VsaEngine.CreateEngine()).ToString()
    End Function

    Private Sub SetTextboxText(wantedWindowHandle As IntPtr, answer As String)
        Dim textboxHandles As List(Of IntPtr) = (From itm In New WindowsEnumerator().GetChildWindows(WantedWindowHandle) Where itm.MainWindowTitle = "" AndAlso itm.ClassName = "Edit" Select itm.hWnd).Select(Function(dummy) CType(dummy, IntPtr)).ToList()
        SendMessageSTRING(textboxHandles(1), WM_SETTEXT, Answer.ToString.Length, Answer.ToString)
    End Sub

    Private Sub ClickButton(wantedWindowHandle As IntPtr)
        For Each itm In From itm1 In New WindowsEnumerator().GetChildWindows(WantedWindowHandle) Where itm1.MainWindowTitle.Contains("Continue")
            SendMessage(itm.hWnd, BM_CLICK, 0, 0)
            Exit Sub
        Next
    End Sub

End Module



