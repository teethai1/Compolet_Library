Imports CompoletAdapter.CompoletClass
Public Class Form1
    Dim IotCompolet As CompoletAdapter.CompoletClass
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox(IotCompolet.Connect("10.1.1.91", "2"))
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox(IotCompolet.IsConnect())
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        IotCompolet = New CompoletAdapter.CompoletClass
        IotCompolet.SetDirectoryPath(My.Application.Info.DirectoryPath & "\AA.txt")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MsgBox(IotCompolet.HeartBeat())
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        MsgBox(IotCompolet.SendEndCleaning())
    End Sub
End Class
