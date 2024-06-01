Imports System.Data.OleDb
Imports System.IO
Public Class Form4
    Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Administrator\Downloads\DictionaryFinal.mdb"
    Dim dataTable As New DataTable
    Dim currentLetter As String = ""
    Dim currentSearch As String = "TextBox1.TextChanged"
    Sub switchPanel(ByVal panel As Form)
        Panel1.Controls.Clear()
        panel.TopLevel = False
        Panel1.Controls.Add(Form6)
        Form6.Dock = DockStyle.Fill
        Form6.Show()
    End Sub
    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        LoadDatabaseData()
        CreateLetterButtons()
        FilterDataBySearch(currentSearch)
        LoadTerms()
    End Sub
    Public Sub LoadDatabaseData()
        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim command As New OleDbCommand("SELECT Terms FROM Dictionary", connection)
            Dim adapter As New OleDbDataAdapter(command)
            adapter.Fill(dataTable)
            UpdateTermsListView(dataTable)
        End Using
        FilterDataByLetter("")
        FilterDataBySearch("")
    End Sub
    Public Sub CreateLetterButtons()
        Dim buttonWidth As Integer = 30
        Dim buttonHeight As Integer = 30
        Dim startingX As Integer = 65
        Dim startingY As Integer = 208
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
    Public Sub LoadTerms()
        Dim connection As New OleDbConnection(connectionString)
        connection.Open()
        Dim query As String = "SELECT Terms FROM Dictionary"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader
        LstTerms1.Items.Clear()
        While reader.Read()
            LstTerms1.Items.Add(reader("Terms").ToString)
            LstTerms1.View = View.Details
            LstTerms1.Columns(0).Width = 310
            LstTerms1.ForeColor = Color.FromArgb(224, 224, 224)
            LstTerms1.Font = New Font("Franklin Gothic Heavy", 20)
        End While
        reader.Close()
        connection.Close()
    End Sub
    Public Sub FilterDataByLetter(ByVal letter As String)
        currentLetter = letter
        Dim filteredData As New DataView(dataTable)
        filteredData.RowFilter = "Terms LIKE '" & letter & "%'"
        LstTerms1.Items.Clear()
        For Each row As DataRowView In filteredData
            LstTerms1.Items.Add(row("Terms").ToString)
        Next
    End Sub
    Public Sub FilterDataBySearch(ByVal word As String)
        currentSearch = word
        Dim filteredData As New DataView(dataTable)
        filteredData.RowFilter = "Terms LIKE '" & word & "%'"
        LstTerms1.Items.Clear()
        For Each row As DataRowView In filteredData
            LstTerms1.Items.Add(row("Terms").ToString)
        Next
    End Sub
    Public Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim clickedButton As Button = CType(sender, Button)
        Dim selectedLetter As String = clickedButton.Text
        FilterDataByLetter(selectedLetter)
    End Sub
    Public Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstTerms1.SelectedIndexChanged
        If LstTerms1.SelectedIndices.Count > 0 Then
            Dim selectedIndex As Integer = LstTerms1.SelectedIndices(0)
            Dim selectedTerm As String = LstTerms1.Items(selectedIndex).Text
            Label2.Text = selectedTerm
            Label2.ForeColor = Color.FromArgb(224, 224, 224)
            ShowDescription(selectedTerm)
        End If
    End Sub
    Public Sub ShowDescription(ByVal Terms As String)
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
        Label3.Text = (description)
        Label3.Font = New Font("Franklin Gothic Heavy", 15)
        Label3.ForeColor = Color.FromArgb(224, 224, 224)
        Label3.Font = AutoSizeFont(Label1.Text, Label1.Width, Label1.Font)
    End Sub
    Public Function AutoSizeFont(ByVal data As String, ByVal maxWidth As Integer, ByVal font As Font) As Font
        Dim textSize As Size = TextRenderer.MeasureText(data, font)
        Dim scaleFactor As Double = Math.Min(1.0, CDbl(maxWidth) / CDbl(textSize.Width))
        Dim newSize As Single
        newSize = Convert.ToSingle(font.Size * scaleFactor)
        newSize = Math.Max(newSize, 12)
        Return New Font(font.FontFamily, newSize, font.Style)
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel1.Visible = True
        Form6.Visible = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        switchPanel(Form6)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim searchTerm As String = TextBox1.Text
        Dim filteredData As New DataView(dataTable)
        filteredData.RowFilter = "Terms LIKE '%" & searchTerm & "%'"
        UpdateTermsListView(filteredData.ToTable)
    End Sub
    Private Sub UpdateTermsListView(ByVal data As DataTable)
        LstTerms1.Items.Clear()
        For Each Row As DataRow In data.Rows
            LstTerms1.Items.Add(New ListViewItem(Row("Terms").ToString))
        Next
    End Sub
End Class
