Imports System.Data.SqlClient
Public Class frmInventoryCategoryRecord
    Dim rdr As SqlDataReader = Nothing
    Dim con As SqlConnection = Nothing
    Dim cmd As SqlCommand = Nothing
    Dim cs As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"
    Private ReadOnly Property Connection() As SqlConnection
        Get
            Dim ConnectionToFetch As New SqlConnection(cs)
            ConnectionToFetch.Open()
            Return ConnectionToFetch
        End Get
    End Property
    Public Function GetData() As DataView
        Dim SelectQry = "SELECT rtrim(CategoryID)[Category ID],rtrim(CategoryName)[Category Name] from InventoryCategory order by CategoryID"

        Dim SampleSource As New DataSet
        Dim TableView As DataView
        Try
            Dim SampleCommand As New SqlCommand()
            Dim SampleDataAdapter = New SqlDataAdapter()
            SampleCommand.CommandText = SelectQry
            SampleCommand.Connection = Connection
            SampleDataAdapter.SelectCommand = SampleCommand
            SampleDataAdapter.Fill(SampleSource)
            TableView = SampleSource.Tables(0).DefaultView
        Catch ex As Exception
            Throw ex
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return TableView
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.DataSource = GetData()
    End Sub

    Private Sub frmInventoryCategoryRecord_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        frmInventoryCategory.Show()
    End Sub

    Private Sub frmInventoryCategoryRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmInventoryCategory.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmInventoryCategory.txtCategoryID.Text = dr.Cells(0).Value.ToString()
            frmInventoryCategory.txtCategoryName.Text = dr.Cells(1).Value.ToString()
            frmInventoryCategory.btnUpdate_Record.Enabled = True
            frmInventoryCategory.btnDelete.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmInventoryCategory.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmInventoryCategory.txtCategoryID.Text = dr.Cells(0).Value.ToString()
            frmInventoryCategory.txtCategoryName.Text = dr.Cells(1).Value.ToString()
            frmInventoryCategory.btnUpdate_record.Enabled = True
            frmInventoryCategory.btnDelete.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class