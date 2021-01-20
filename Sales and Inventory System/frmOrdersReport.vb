Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmOrdersReport
    Dim dtable As DataTable
    Dim adp As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim cs As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        dtpOrderDateFrom.Text = Today
        dtpOrderDateFrom.Text = Today
        CrystalReportViewer1.ReportSource = Nothing
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        cmbStatus.Text = ""
        CrystalReportViewer2.ReportSource = Nothing
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        cmbCustomerName.Text = ""
        CrystalReportViewer3.ReportSource = Nothing
    End Sub

   

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        
        cmbCustomerName.Text = ""
        CrystalReportViewer3.ReportSource = Nothing
        dtpOrderDateFrom.Text = Today
        dtpOrderDateFrom.Text = Today
        CrystalReportViewer1.ReportSource = Nothing
        cmbStatus.Text = ""
        CrystalReportViewer2.ReportSource = Nothing
    End Sub
    
    Sub fillCustomerName()

        Try

            Dim CN As New SqlConnection(cs)

            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  RTRIM(CustomerName) FROM orderinfo", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCustomerName.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbCustomerName.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmOrdersReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub
    Private Sub frmOrdersReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCustomerName()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New OrderInfo_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where OrderDate between @date1 and @date2 order by OrderDate"
            MyCommand.Parameters.Add("@date1", SqlDbType.DateTime, 30, "BillingDate").Value = dtpOrderDateFrom.Value.Date
            MyCommand.Parameters.Add("@date2", SqlDbType.DateTime, 30, "BillingDate").Value = dtpOrderDateTo.Value.Date

            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try

            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New OrderInfo_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where OrderStatus = '" & cmbStatus.Text & "' order by OrderDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer2.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try

            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New OrderInfo_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where CustomerName = '" & cmbCustomerName.Text & "' order by OrderDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer3.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class