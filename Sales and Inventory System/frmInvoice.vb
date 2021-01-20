Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmInvoice

    Private Sub frmInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Dim rpt As New rptInvoice() 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New SqlConnection("Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * FROM  BillInfo INNER JOIN ProductSold ON BillInfo.InvoiceNo = ProductSold.InvoiceNo INNER JOIN Customer ON BillInfo.CustomerNo = Customer.CustomerNo WHERE billinfo.invoiceno= '" & frmInVoiceNo.cmbInvoiceNo.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            myDA.Fill(myDS, "Customer")
            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class