Imports System.Data.OleDb

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
        Try
            If CheckOleDbconnection() Then

                oledb_connection.Open()
                sql_string = "INSERT INTO tblitems(ITEMNAME,ITEMDESCRIPTION,QTY,PRICE) VALUES ('" & txt_itemname.Text & "','" & txt_itemdescription.Text & "','" & Val(txt_qty.Text) & "','" & Val(txt_price.Text) & "');"
                oledb_command.Connection = oledb_connection
                oledb_command.CommandText = sql_string
                Dim i As Integer = oledb_command.ExecuteNonQuery

                If i > 0 Then
                    MsgBox("New Item has been saved!")
                    ClearTable()
                    PopulationTableItems()
                    ClearInputFields()
                Else
                    MsgBox("Cannot Save Item!!")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oledb_connection.Close()
        End Try

    End Sub

    Private Sub dg_view_item_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg_view_item.CellContentClick
        RowSelect()
    End Sub

    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click
        Try
            If CheckOleDbconnection() Then
                oledb_connection.Open()
                sql_string = "UPDATE tblitems SET ITEMNAME='" & txt_itemname.Text & "',ITEMDESCRIPTION='" & txt_itemdescription.Text & "',QTY=" & Val(txt_qty.Text) & ",PRICE=" & Val(txt_price.Text) & " WHERE ID=" & Val(Me.Text) & ""
                oledb_command.Connection = oledb_connection
                oledb_command.CommandText = sql_string
                Dim i As Integer = oledb_command.ExecuteNonQuery
                If i > 0 Then
                    MsgBox("ID No." & Me.Text & " has been UPDATED!")
                    ClearTable()
                    PopulationTableItems()
                    ClearInputFields()
                Else
                    MsgBox("ID No." & Me.Text & "CANNOT UPDATE")
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oledb_connection.Close()
        End Try

    End Sub

    Private Sub dg_view_item_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg_view_item.CellClick
        RowSelect()
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        Try
            If CheckOleDbconnection() Then
                oledb_connection.Open()
                sql_string = "DELETE  * FROM tblitems WHERE ID=" & Val(Me.Text) & ""
                oledb_command.Connection = oledb_connection
                oledb_command.CommandText = sql_string
                Dim i As Integer = oledb_command.ExecuteNonQuery
                If i > 0 Then
                    MsgBox("ID No." & Me.Text & " has been DELETED!")
                    ClearTable()
                    PopulationTableItems()
                Else
                    MsgBox("ID No." & Me.Text & "CANNOT DELETE")
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oledb_connection.Close()
        End Try
    End Sub
End Class
