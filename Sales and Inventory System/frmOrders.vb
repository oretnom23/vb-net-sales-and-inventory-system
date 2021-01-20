Imports System.Data.SqlClient
Public Class frmOrders
    Dim rdr As SqlDataReader = Nothing
    Dim dtable As DataTable
    Dim con As SqlConnection = Nothing
    Dim adp As SqlDataAdapter
    Dim ds As DataSet
    Dim cmd As SqlCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try


            con = New SqlConnection(cs)

            con.Open()
            cmd = New SqlCommand("SELECT rtrim(OrderNo)[Order No.],rtrim(OrderDate)[Order Date],rtrim(OrderStatus)[OrderStatus],rtrim(CustomerNo)[Distributor ID],rtrim(CustomerName)[Customer Name],rtrim(TotalAmount)[Total Amount] from Orderinfo where OrderStatus='Uncompleted' and OrderDate between @date1 and @date2 order by OrderDate", con)
            cmd.Parameters.Add("@date1", SqlDbType.DateTime, 30, "OrderDate").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@date2", SqlDbType.DateTime, 30, "OrderDate").Value = DateTimePicker1.Value.Date


            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)

            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "OrderInfo")

            DataGridView1.DataSource = myDataSet.Tables("Orderinfo").DefaultView
          
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmOrders_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtOrderNo.Text = dr.Cells(0).Value.ToString()


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class