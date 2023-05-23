Imports System.Data.OleDb
Imports Microsoft.Reporting.WinForms

Public Class Form1
    'OLE DB INITIAL COLLECTION
    Private oledb_connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Admin\Documents\inventorydb.accdb")

    Private sql_string As String
    Private oledb_command As New OleDbCommand
    Private data_table As New DataTable
    Private oledb_data_adapter As New OleDb.OleDbDataAdapter

    Function CheckOleDbconnection() As Boolean
        sql_string = ""
        Return oledb_connection.State <> ConnectionState.Open
    End Function

    Function ClearTable()
        dg_view_item.DataSource = Nothing
        data_table.Clear()
        dg_view_item.Refresh()
    End Function

    Function ClearInputFields()
        txt_itemname.Text = ""
        txt_itemdescription.Text = ""
        txt_price.Text = ""
        txt_qty.Text = ""
    End Function

    Function RowSelect()

        Me.Text = dg_view_item.CurrentRow.Cells(0).Value
        txt_itemname.Text = dg_view_item.CurrentRow.Cells(1).Value
        txt_itemdescription.Text = dg_view_item.CurrentRow.Cells(2).Value
        txt_qty.Text = dg_view_item.CurrentRow.Cells(3).Value
        txt_price.Text = dg_view_item.CurrentRow.Cells(4).Value


    End Function



    Function OleDbComm(ByVal string_qry As String, ByVal db_con As OleDbConnection)
        Try
            oledb_connection.Open()
            oledb_command.Connection = db_con
            oledb_command.CommandText = string_qry
            Return oledb_command.ExecuteNonQuery

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oledb_connection.Close()
        End Try

    End Function

    Function PopulationTableItems()
        Try
            If CheckOleDbconnection() Then
                oledb_connection.Open()
            End If

            If oledb_connection.State = ConnectionState.Open Then

                sql_string = "SELECT * FROM tblitems"
                oledb_command.Connection = oledb_connection
                oledb_command.CommandText = sql_string
                oledb_data_adapter.SelectCommand = oledb_command
                oledb_data_adapter.Fill(data_table)
                dg_view_item.DataSource = data_table
            Else
                MessageBox.Show("Database Not Connected")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            oledb_connection.Close()
        End Try
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulationTableItems()
    End Sub


    Private Sub btn_save_Click(sender As Object, e As EventArgs) Handles btn_save.Click

        If CheckOleDbconnection() Then
            If (OleDbComm("INSERT INTO tblitems(ITEMNAME,ITEMDESCRIPTION,QTY,PRICE) VALUES ('" & txt_itemname.Text & "','" & txt_itemdescription.Text & "','" & Val(txt_qty.Text) & "','" & Val(txt_price.Text) & "');", oledb_connection) > 0) Then
                MsgBox("Item has been SAVED!")
                ClearTable()
                PopulationTableItems()
                ClearInputFields()
            Else
                MsgBox("Item not save!")
            End If
        End If


    End Sub

    Private Sub dg_view_item_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg_view_item.CellContentClick
        RowSelect()
    End Sub

    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click

        If CheckOleDbconnection() Then
            If (OleDbComm("UPDATE tblitems SET ITEMNAME='" & txt_itemname.Text & "',ITEMDESCRIPTION='" & txt_itemdescription.Text & "',QTY=" & Val(txt_qty.Text) & ",PRICE=" & Val(txt_price.Text) & " WHERE ID=" & Val(Me.Text) & "", oledb_connection) > 0) Then
                MsgBox("Item have been UPDATED!")
                ClearTable()
                PopulationTableItems()
                ClearInputFields()
            Else
                MsgBox("Item not UPDATE!")
            End If
        End If

    End Sub

    Private Sub dg_view_item_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg_view_item.CellClick
        RowSelect()
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click


        If CheckOleDbconnection() Then
            If (OleDbComm("DELETE  * FROM tblitems WHERE ID=" & Val(Me.Text) & "", oledb_connection) > 0) Then
                MsgBox("Item have been DELETED!")
                ClearTable()
                PopulationTableItems()
                ClearInputFields()
            Else
                MsgBox("Item not deleted!")
            End If
        End If


    End Sub

    Private Sub btn_ge_report_Click(sender As Object, e As EventArgs) Handles btn_ge_report.Click
        Dim oledb_reader As OleDbDataReader
        Dim data_set As DataSet = New inventorydbDataSet


        Report_Form.Show()


        Try

            sql_string = "SELECT * FROM tblitems"
            oledb_command = New OleDbCommand(sql_string, oledb_connection)
            oledb_command.CommandType = CommandType.Text
            oledb_command.Connection.Open()
            oledb_reader = oledb_command.ExecuteReader
            data_set.Tables(0).Clear()
            data_set.Tables(0).Load(oledb_reader)
            oledb_reader.Close()
            oledb_connection.Close()
            Dim rpt_data_source = New ReportDataSource("dataset_inventory_report", data_set.Tables(0))
            rpt_data_source.Name = "dataset_inventory_report"
            rpt_data_source.Value = data_set.Tables(0)




            With Report_Form.report_viewer
                .ProcessingMode = ProcessingMode.Local
                .LocalReport.DataSources.Clear()
                .LocalReport.DataSources.Add(rpt_data_source)
                .LocalReport.Refresh()
                .PrinterSettings.Copies = 1
                '.SetDisplayMode(DisplayMode.PrintLayout)
                .ShowRefreshButton = True
                .Text = "Inventory Report"
                .RefreshReport()


            End With




        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            'oledb_connection.Close()
        End Try

    End Sub

End Class
