Imports System.Data.OleDb
Imports System.Drawing.Font
Public Class Form1
    Dim Connect As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Documents\Database1.mdb")
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Dim iExit As DialogResult
            iExit = MessageBox.Show("Confirm if you want to Exit", " Login System", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

            If (iExit = DialogResult.Yes) Then Application.Exit()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmd As OleDbCommand
        Dim dr As OleDbDataReader
        Dim checker As Integer
        Dim username As String = TextBox1.Text.Trim()
        Dim password As String = TextBox2.Text.Trim()
        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Documents\Database1.mdb"
        Dim selectCommand As String = "SELECT * FROM Table1 WHERE Username = @Username AND Password = @Password"
        Dim minLength As Integer = 5
        Dim maxLength As Integer = 12
        Button1.FlatStyle = FlatStyle.Flat
        Button1.FlatAppearance.BorderSize = 0
        If TextBox1.TextLength < minLength Then
            TextBox1.ForeColor = Color.Red
            MessageBox.Show("Username must be at least 5-8 characters long.", "Invalid Username", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox1.TextLength > 8 Then
            TextBox1.ForeColor = Color.Red
            MessageBox.Show("Username must be at least 5-8 characters long.", "Invalid Username", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.TextLength < 8 Then
            TextBox2.ForeColor = Color.Red
            MessageBox.Show("Password must be at least 10 characters long and contains uppercase, lowercase letters, and numbers.", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.TextLength > maxLength Then
            TextBox2.ForeColor = Color.Red
            MessageBox.Show("Password must be at least 10 characters long and contains uppercase, lowercase letters, and numbers.", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            TextBox1.ForeColor = Color.Silver
            TextBox2.ForeColor = Color.Silver
        End If
        Try
            Using connection As New OleDbConnection(connectionString)
                Connect.Open()
                Using command As New OleDbCommand(selectCommand, connection)
                    command.Parameters.AddWithValue("@Username", username)
                    command.Parameters.AddWithValue("@Password", password)
                    cmd = Connect.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select * from Table1 where username = '" + TextBox1.Text + "' and password = '" + TextBox2.Text + "'"

                    dr = cmd.ExecuteReader()
                    checker = 0

                    Do While (dr.Read())
                        checker = checker + 1
                    Loop
                    If dr.HasRows Then
                        Me.Hide()
                        Form2.Show()
                    Else
                        MessageBox.Show("Invalid username or password. Please Try Again.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Connect.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Form3.Show()
    End Sub
    Private Sub Textbox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        Label3.Visible = TextBox1.Text = ""
        Label4.Visible = TextBox2.Text = ""
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        TextBox1.Focus()
    End Sub
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        TextBox2.Focus()
    End Sub
    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        PictureBox2.Visible = False
        PictureBox3.Visible = True
        AddHandler PictureBox3.Click, AddressOf PictureBox3_Click
        TextBox2.PasswordChar = vbNullChar
    End Sub
    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        PictureBox3.Visible = False
        PictureBox2.Visible = True
        AddHandler PictureBox2.Click, AddressOf PictureBox2_Click
        TextBox2.PasswordChar = "*"
    End Sub
End Class
