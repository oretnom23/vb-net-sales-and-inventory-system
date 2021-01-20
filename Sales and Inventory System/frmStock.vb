Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text
Public Class frmStock
    Dim rdr As SqlDataReader = Nothing
    Dim dtable As DataTable
    Dim con As SqlConnection = Nothing
    Dim adp As SqlDataAdapter
    Dim ds As DataSet
    Dim cmd As SqlCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"

    Sub clear()
        txtStockID.Text = ""
        txtCartons.Text = ""
        txtCategory.Text = ""
        txtPackets.Text = ""
        txtWeight.Text = ""
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtTotalPackets.Text = ""
        dtpStockDate.Text = Today
        Button2.Focus()
    End Sub
    Private Const ConnectionString As String = "Data Source=.\SqlExpress; Integrated Security=True; AttachDbFilename=|DataDirectory|\SI_DB.mdf; User Instance=true;"
    Private ReadOnly Property Connection() As SqlConnection
        Get
            Dim ConnectionToFetch As New SqlConnection(ConnectionString)
            ConnectionToFetch.Open()
            Return ConnectionToFetch
        End Get
    End Property
    Public Function GetData() As DataView
        Dim SelectQry = "SELECT RTRIM(ProductCode)[Product Code],RTRIM(ProductName)[Product Name],RTRIM(Weight)[Weight],sum(Cartons)[Cartons],Sum(TotalPackets)[Total Packets] FROM stock where Cartons > 0 and TotalPackets > 0   group by ProductCode,ProductName,Weight order by ProductName "
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

    Private Sub txtPackets_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPackets.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub
    Private Sub txtPackets_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPackets.TextChanged
        txtTotalPackets.Text = CInt(Val(txtCartons.Text) * Val(txtPackets.Text))
    End Sub

    Private Sub txtCartons_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCartons.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub

    Private Sub txtCartons_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCartons.TextChanged
        txtTotalPackets.Text = CInt(Val(txtCartons.Text) * Val(txtPackets.Text))
    End Sub


    Private Sub NewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        clear()
        Save.Enabled = True
        Update_Record.Enabled = False
        Delete.Enabled = False
    End Sub
    Public Shared Function GetUniqueKey(ByVal maxSize As Integer) As String
        Dim chars As Char() = New Char(61) {}
        chars = "123456789".ToCharArray()
        Dim data As Byte() = New Byte(0) {}
        Dim crypto As New RNGCryptoServiceProvider()
        crypto.GetNonZeroBytes(data)
        data = New Byte(maxSize - 1) {}
        crypto.GetNonZeroBytes(data)
        Dim result As New StringBuilder(maxSize)
        For Each b As Byte In data
            result.Append(chars(b Mod (chars.Length)))
        Next
        Return result.ToString()
    End Function
    Sub auto()
        txtStockID.Text = "ST-" & GetUniqueKey(6)
    End Sub
    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        If Len(Trim(txtProductCode.Text)) = 0 Then
            MessageBox.Show("Please select product code", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProductCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProductName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(txtWeight.Text)) = 0 Then
            MessageBox.Show("Please select weight", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtWeight.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCartons.Text)) = 0 Then
            MessageBox.Show("Please enter cartons", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCartons.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPackets.Text)) = 0 Then
            MessageBox.Show("Please enter packets", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPackets.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select ProductCode from stock where ProductCode=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.NChar, 20, "ProductCode"))
            cmd.Parameters("@find").Value = txtProductCode.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Entry for product already exists" & vbCrLf & "You can not make duplicate entry" & vbCrLf & "for the same product" & vbCrLf & "please update the stock of product", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If


            auto()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct1 As String = "select stockid from stock where stockid=@find"

            cmd = New SqlCommand(ct1)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.NChar, 20, "stockid"))
            cmd.Parameters("@find").Value = txtStockID.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Stock ID Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

            Else



                con = New SqlConnection(cs)
                con.Open()

                Dim cb As String = "insert into stock(StockID,productcode,productname,category,weight,stockdate,Cartons,Packets,TotalPackets) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9)"

                cmd = New SqlCommand(cb)

                cmd.Connection = con
                cmd.Parameters.Add(New SqlParameter("@d1", System.Data.SqlDbType.NChar, 20, "StockID"))

                cmd.Parameters.Add(New SqlParameter("@d2", System.Data.SqlDbType.NChar, 20, "productcode"))
                cmd.Parameters.Add(New SqlParameter("@d3", System.Data.SqlDbType.VarChar, 250, "productname"))
                cmd.Parameters.Add(New SqlParameter("@d4", System.Data.SqlDbType.VarChar, 150, "category"))
                cmd.Parameters.Add(New SqlParameter("@d5", System.Data.SqlDbType.NChar, 10, "weight"))
                cmd.Parameters.Add(New SqlParameter("@d6", System.Data.SqlDbType.NChar, 30, "Stockdate"))
                cmd.Parameters.Add(New SqlParameter("@d7", System.Data.SqlDbType.Int, 250, "cartons"))
                cmd.Parameters.Add(New SqlParameter("@d8", System.Data.SqlDbType.Int, 10, "Packets"))
                cmd.Parameters.Add(New SqlParameter("@d9", System.Data.SqlDbType.Int, 10, "TotalPackets"))

                cmd.Parameters("@d1").Value = txtStockID.Text
                cmd.Parameters("@d2").Value = txtProductCode.Text

                cmd.Parameters("@d3").Value = txtProductName.Text


                cmd.Parameters("@d4").Value = txtCategory.Text

                cmd.Parameters("@d5").Value = txtWeight.Text
                cmd.Parameters("@d6").Value = dtpStockDate.Text

                cmd.Parameters("@d7").Value = CInt(txtCartons.Text)
                cmd.Parameters("@d8").Value = CInt(txtPackets.Text)

                cmd.Parameters("@d9").Value = CInt(txtTotalPackets.Text)
                cmd.Parameters("@d8").Value = CInt(txtPackets.Text)

                cmd.Parameters("@d9").Value = CInt(txtTotalPackets.Text)

                cmd.ExecuteReader()
                MessageBox.Show("Successfully saved", "Stock Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Save.Enabled = False
                DataGridView1.DataSource = GetData()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                con.Close()
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmStock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub

    Private Sub frmStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = GetData()
    End Sub

    Private Sub Update_Record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_Record.Click
        Try
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "update stock set productcode=@d2,productname=@d3,category=@d4,weight=@d5,stockdate=@d6,Cartons=@d7,Packets=@d8,TotalPackets=@d9 where stockid=@d1"

            cmd = New SqlCommand(cb)

            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@d1", System.Data.SqlDbType.NChar, 20, "StockID"))

            cmd.Parameters.Add(New SqlParameter("@d2", System.Data.SqlDbType.NChar, 20, "productcode"))
            cmd.Parameters.Add(New SqlParameter("@d3", System.Data.SqlDbType.VarChar, 250, "productname"))
            cmd.Parameters.Add(New SqlParameter("@d4", System.Data.SqlDbType.VarChar, 150, "category"))
            cmd.Parameters.Add(New SqlParameter("@d5", System.Data.SqlDbType.NChar, 10, "weight"))
            cmd.Parameters.Add(New SqlParameter("@d6", System.Data.SqlDbType.NChar, 30, "Stockdate"))
            cmd.Parameters.Add(New SqlParameter("@d7", System.Data.SqlDbType.Int, 250, "cartons"))
            cmd.Parameters.Add(New SqlParameter("@d8", System.Data.SqlDbType.Int, 10, "Packets"))
            cmd.Parameters.Add(New SqlParameter("@d9", System.Data.SqlDbType.Int, 10, "TotalPackets"))

            cmd.Parameters("@d1").Value = txtStockID.Text
            cmd.Parameters("@d2").Value = txtProductCode.Text

            cmd.Parameters("@d3").Value = txtProductName.Text


            cmd.Parameters("@d4").Value = txtCategory.Text

            cmd.Parameters("@d5").Value = txtWeight.Text
            cmd.Parameters("@d6").Value = dtpStockDate.Text

            cmd.Parameters("@d7").Value = CInt(txtCartons.Text)
            cmd.Parameters("@d8").Value = CInt(txtPackets.Text)

            cmd.Parameters("@d9").Value = CInt(txtTotalPackets.Text)


            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Stock Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Update_Record.Enabled = False
            DataGridView1.DataSource = GetData()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete.Click
        Try



            If MessageBox.Show("Do you really want to delete the record?", "Stock Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try



            Dim RowsAffected As Integer = 0

         
            con = New SqlConnection(cs)

            con.Open()


            Dim cq As String = "delete from stock where stockid=@DELETE1;"


            cmd = New SqlCommand(cq)

            cmd.Connection = con

            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.NChar, 20, "stockid"))


            cmd.Parameters("@DELETE1").Value = Trim(txtStockID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = GetData()
                clear()

                Update_Record.Enabled = False
                Delete.Enabled = False
              
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = GetData()
                clear()
                Update_Record.Enabled = False
                Delete.Enabled = False

                If con.State = ConnectionState.Open Then

                    con.Close()
                End If

                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.clear()
        frmProductsRecord.DataGridView4.DataSource = Nothing
        frmProductsRecord.cmbCategory.Text = ""
        frmProductsRecord.cmbWeight.Text = ""
        frmProductsRecord.DataGridView3.DataSource = Nothing
        frmProductsRecord.cmbProductName.Text = ""
        frmProductsRecord.txtProduct.Text = ""
        frmProductsRecord.DataGridView2.DataSource = Nothing
        frmProductsRecord.DataGridView1.DataSource = Nothing
        frmProductsRecord.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.clear()
       
        frmStockDetails1.cmbProductName.Text = ""
        frmStockDetails1.txtProduct.Text = ""
        frmStockDetails1.DataGridView2.DataSource = Nothing
        frmStockDetails1.cmbCategory.Text = ""
        frmStockDetails1.DataGridView3.DataSource = Nothing
        frmStockDetails1.cmbWeight.Text = ""
        frmStockDetails1.DataGridView4.DataSource = Nothing

        frmStockDetails1.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        System.Diagnostics.Process.Start("Calc.exe")
    End Sub
End Class