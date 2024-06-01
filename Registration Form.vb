Imports System.Data.OleDb
Imports System.IO
Public Class Form3
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim minLength As Integer = 2
        Dim maxLength As Integer = 20
        Dim FirstName As String = TextBox1.Text.Trim()
        Dim LastName As String = TextBox2.Text.Trim()
        Dim StudentNumber As String = TextBox3.Text.Trim()
        Dim Email As String = TextBox4.Text.Trim()
        Dim Username As String = TextBox5.Text.Trim()
        Dim Password As String = TextBox6.Text.Trim()
        If Password.Length <> 10 OrElse Not Password.Any(Function(c) Char.IsUpper(c)) OrElse Not Password.Any(Function(c) Char.IsLower(c)) OrElse Not Password.Any(Function(c) Char.IsNumber(c)) Then
            TextBox6.ForeColor = Color.Red
            MessageBox.Show("Password must be at least 10 characters long and contains uppercase, lowercase letters, and numbers.", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf FirstName.Length < 2 And FirstName.Length > 12 Then
            TextBox1.ForeColor = Color.Red
            MessageBox.Show("First Name must be at least 2-12 characters long.", "Invalid First Name", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf LastName.Length < 2 And LastName.Length > 12 Then
            TextBox1.ForeColor = Color.Red
            MessageBox.Show("Last Name must be at least 2-12 characters long.", "Invalid Last Name", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf StudentNumber.Length <> 10 Then
            MessageBox.Show("Student Number must be at least 10 characters long.", "Invalid Student Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf Email.Length < 10 And Email.Length > 20 Then
            MessageBox.Show("Email must be at least 10-20 characters long.", "Invalid Email", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf Username.Length < 6 And Username.Length > 8 Then
            MessageBox.Show("Username must be at least 6-8 characters long.", "Invalid Username", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Else
            TextBox6.ForeColor = Color.Silver
        End If
        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Documents\Database1.mdb"
        Dim insertCommand As String = "INSERT INTO Table1 ([First Name], [Last Name], [Student Number], [Email], [Username], [Password]) VALUES (@TextBox1.Text, @TextBox2.Text, @TextBox3.Text, @TextBox4.Text, @Text.Box5.Text, @TextBox6.Text)"
        Using Connection As New OleDbConnection(connectionString)
                Connection.Open()
                Using Command As New OleDbCommand(insertCommand, Connection)
                Command.Parameters.AddWithValue("@FirstName", FirstName)
                    Command.Parameters.AddWithValue("@LastName", LastName)
                    Command.Parameters.AddWithValue("@StudentNumber", StudentNumber)
                    Command.Parameters.AddWithValue("@Email", Email)
                    Command.Parameters.AddWithValue("@Username", Username)
                Command.Parameters.AddWithValue("@Password", Password)
                Try
                    Command.ExecuteNonQuery()
                    MessageBox.Show("Registration Complete!", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message & vbCrLf & "Please contact the administrator.", "Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                End Using            
        End Using
        Me.Hide()
        Form1.Show()
    End Sub
    Function isValidEmail(ByVal email As String) As Boolean
        Return email.Contains("@") AndAlso email.Contains(".")
    End Function
    Private Function GetImage(ByVal imageName As String) As Image
        Return Image.FromFile(imageName)
    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Label2.Visible = TextBox1.Text = ""
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Label3.Visible = TextBox2.Text = ""
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        Label9.Visible = TextBox3.Text = ""
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Label8.Visible = TextBox4.Text = ""
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Label11.Visible = TextBox5.Text = ""
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Label10.Visible = TextBox6.Text = ""
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        TextBox1.Focus()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        TextBox2.Focus()
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        TextBox3.Focus()
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        TextBox4.Focus()
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        TextBox5.Focus()
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        TextBox6.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim iExit As DialogResult
            iExit = MessageBox.Show("Confirm if you want to Cancel", " Login System", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

            If (iExit = DialogResult.Yes) Then Application.Exit()
            Form1.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private images As String() = {"C:\Users\Administrator\Documents\Visual Studio 2010\Projects\LoginFormDatabase\LoginFormDatabase\Resources\COMDIC.png", "C:\Users\Administrator\Documents\Visual Studio 2010\Projects\LoginFormDatabase\LoginFormDatabase\Resources\4c2c7f37-8ce4-4a70-911c-a5135b0d6856-removebg-preview.png"}
    Private currentImageIndex As Integer = 0
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PictureBox1.Image = GetImage(images(0))
        Timer1.Enabled = True
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        currentImageIndex = (currentImageIndex + 1) Mod images.Length
        PictureBox1.Image = GetImage(images(currentImageIndex))
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        Me.Hide()
        Form1.Show()
    End Sub
End Class
