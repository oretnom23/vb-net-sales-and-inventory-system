Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmSalesReport
    Dim dtable As DataTable
    Dim adp As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim cs As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            Dim rpt As New rptSales() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Billinfo where BillingDate between @date1 and @date2 order by BillingDate"
            MyCommand.Parameters.Add("@date1", SqlDbType.DateTime, 30, "BillingDate").Value = dtpInvoiceDateFrom.Value.Date
            MyCommand.Parameters.Add("@date2", SqlDbType.DateTime, 30, "BillingDate").Value = dtpInvoiceDateTo.Value.Date

            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Billinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        CrystalReportViewer1.ReportSource = Nothing
        dtpInvoiceDateFrom.Text = Today
        dtpInvoiceDateTo.Text = Today
        CrystalReportViewer2.ReportSource = Nothing
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        CrystalReportViewer3.ReportSource = Nothing
       
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        dtpInvoiceDateFrom.Text = Today
        dtpInvoiceDateTo.Text = Today
        CrystalReportViewer1.ReportSource = Nothing
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        CrystalReportViewer2.ReportSource = Nothing
        cmbCustomerName.Text = ""
    End Sub

  

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try

            Dim rpt As New rptSales() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Billinfo where CustomerName= '" & cmbCustomerName.Text & "' order by BillingDate"
           
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Billinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer2.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

  
  
    Sub fillCustomerName()

        Try

            Dim CN As New SqlConnection(cs)

            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  RTRIM(CustomerName) FROM Billinfo", CN)
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

    Private Sub frmSalesReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub
    Private Sub frmSalesReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCustomerName()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        CrystalReportViewer3.ReportSource = Nothing
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            Dim rpt As New rptSales() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Billinfo where PaymentDue > 0 and BillingDate between @date1 and @date2 order by BillingDate"
            MyCommand.Parameters.Add("@date1", SqlDbType.DateTime, 30, "BillingDate").Value = DateTimePicker2.Value.Date
            MyCommand.Parameters.Add("@date2", SqlDbType.DateTime, 30, "BillingDate").Value = DateTimePicker1.Value.Date

            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Billinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer3.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class