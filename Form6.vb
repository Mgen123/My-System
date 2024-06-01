Imports System.Data.OleDb
Public Class Form6
    Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Downloads\DictionaryFinal.mdb"
    Dim dataTable As New DataTable
    Dim currentLetter As String = ""
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDatabaseData()
        CreateLetterButtons()
        LoadTerms()
    End Sub
    Private Sub LoadTerms()
        Dim connection As New OleDbConnection(connectionString)
        connection.Open()
        Dim query As String = "SELECT Terms FROM Dictionary"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader
        ListView1.Items.Clear()
        While reader.Read()
            ListView1.Items.Add(reader("Terms").ToString)
            ListView1.View = View.Details
            ListView1.Columns(0).Width = 310
            ListView1.ForeColor = Color.FromArgb(224, 224, 224)
            ListView1.Font = New Font("Franklin Gothic Heavy", 20)
        End While
        reader.Close()
        connection.Close()
    End Sub
    Private Sub LoadDatabaseData()
        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim command As New OleDbCommand("SELECT Terms FROM Dictionary", connection)
            Dim adapter As New OleDbDataAdapter(command)
            adapter.Fill(dataTable)
        End Using
        FilterDataByLetter("")
    End Sub
    Private Sub CreateLetterButtons()
        Dim buttonWidth As Integer = 30
        Dim buttonHeight As Integer = 30
        Dim startingX As Integer = 66
        Dim startingY As Integer = 26
        Dim buttonMargin As Integer = 5

        For i As Integer = 0 To 25
            Dim btnLetter As New Button
            btnLetter.Text = Chr(65 + i).ToString
            btnLetter.Width = buttonWidth
            btnLetter.Height = buttonHeight
            btnLetter.FlatStyle = FlatStyle.Flat
            btnLetter.ForeColor = Color.WhiteSmoke
            btnLetter.Location = New Point(startingX + (i * (buttonHeight + buttonMargin)), startingY)
            AddHandler btnLetter.Click, AddressOf ButtonClick
            Controls.Add(btnLetter)
        Next
    End Sub
    Private Sub FilterDataByLetter(ByVal letter As String)
        currentLetter = letter
        Dim filteredData As New DataView(dataTable)
        filteredData.RowFilter = "Terms LIKE '" & letter & "%'"
        ListView1.Items.Clear()
        For Each row As DataRowView In filteredData
            ListView1.Items.Add(row("Terms").ToString)
        Next
    End Sub

    Private Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim clickedButton As Button = CType(sender, Button)
        Dim selectedLetter As String = clickedButton.Text
        FilterDataByLetter(selectedLetter)
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedIndices.Count > 0 Then
            Dim selectedIndex As Integer = ListView1.SelectedIndices(0)
            Dim selectedTerm As String = ListView1.Items(selectedIndex).Text
            Label2.Text = selectedTerm
            Label2.ForeColor = Color.FromArgb(224, 224, 224)
            ShowDescription(selectedTerm)
        End If
    End Sub
    Private Sub ShowDescription(ByVal Terms As String)
        Dim description As String = ""
        Dim connection As New OleDbConnection(connectionString)
        connection.Open()
        Dim query As String = "SELECT Description FROM Dictionary WHERE Terms = @Terms"
        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@Terms", Terms)
        Dim reader As OleDbDataReader = command.ExecuteReader
        If reader.HasRows Then
            reader.Read()
            description = reader("description").ToString
        End If
        reader.Close()
        connection.Close()
        Label1.Text = ""
        Label1.Text = (description)
        Label1.Font = New Font("Franklin Gothic Heavy", 15)
        Label1.ForeColor = Color.FromArgb(224, 224, 224)
        Label1.Font = AutoSizeFont(Label1.Text, Label1.Width, Label1.Font)
    End Sub
    Public Function AutoSizeFont(ByVal data As String, ByVal maxWidth As Integer, ByVal font As Font) As Font
        Dim textSize As Size = TextRenderer.MeasureText(data, font)
        Dim scaleFactor As Double = Math.Min(1.0, CDbl(maxWidth) / CDbl(textSize.Width))
        Dim newSize As Single
        newSize = Convert.ToSingle(font.Size * scaleFactor)
        newSize = Math.Max(newSize, 12)
        Return New Font(font.FontFamily, newSize, font.Style)
    End Function
End Class