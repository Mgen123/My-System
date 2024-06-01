Imports System.Data.OleDb
Imports System.IO
Public Class Form2
    Dim Connect As New OledbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Documents\Database1.mdb")
    Dim i As Integer
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Connect.Open()
            Label3.Text = "Connecting..."
            Label3.ForeColor = Color.White
        Catch ex As Exception
            Label3.Text = "Connection Failed"
            Label3.ForeColor = Color.Red
        End Try
        Connect.Close()
        If Label3.Text = "Connecting..." Then
            Timer1.Enabled = True
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label3.Text = "Connected"
        Label3.ForeColor = Color.Lime
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.Close()
        Form4.Show()
    End Sub
    Private Sub Timer3_Tick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If Label3.Text = "Connected" Then
            Label1.Text = "Successful!"
            Label1.ForeColor = Color.Lime
        End If
    End Sub
End Class
