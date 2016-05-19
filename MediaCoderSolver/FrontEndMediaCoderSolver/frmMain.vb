Public Class FrmMain

    'This class is neccessary because ConsoleMediaCoderSolver sometimes stops responding if it fails to execute correctly.

    Private ReadOnly _progPath As String = Application.StartupPath & "\ConsoleMediaCoderSolver.exe"
    Private Sub cbxEnabled_CheckedChanged(sender As Object, e As EventArgs) Handles cbxEnabled.CheckedChanged
        tmrRunApp.Enabled = cbxEnabled.Checked
    End Sub

    Private Sub tmrRunApp_Tick(sender As Object, e As EventArgs) Handles tmrRunApp.Tick
        If IO.File.Exists(_progPath) = False Then
            IO.File.WriteAllBytes(_progPath, My.Resources.ConsoleMediaCoderSolver)
        End If
        For Each p As Process In Process.GetProcessesByName("ConsoleMediaCoderSolver")
            p.Kill()
        Next
        Process.Start(_progPath)
    End Sub
End Class
