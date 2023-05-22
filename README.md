# VB_NET_OLEDB_MSACCESS_APP
sample  CRUD application from VB.NET with MS ACCESS DB

Preparation for DICT Proficieny HandsOn Exam




OLEDB
Connector/ODBC
ADO.NET ODBC


#import  OLEDB package
        Imports System.Data.OleDb

#DECLARE OLEDB CONNECTION from DATABASE  MS ACESS( inventorydb.accdb) , ACE.OLEDB.12 can be upgraded  to 16 and install the ACCESS #DATABASE ENGINE installer from internet
        Dim oledb_connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Admin\Documents\inventorydb.accdb")


#DECLARE for STRING SQL QUERY "eg: SELECT * FROM tableTest where"
    Dim sql_string As String
#DECLARE OLEDB COMMANDS
    Dim oledb_command As New OleDbCommand
#DECLARE DATA TABLE
    Dim data_table As New DataTable
#DECLARE OLEDB DATA ADAPTER
    Dim oledb_data_adapter As New OleDb.OleDbDataAdapter



##################### DISPLAY ALL THE DATA FROM DB to DATAGRID VIEW ######################################################

    #open the oledbconnection from MS ACCESS
        oledb_connection.Open()

    #DECLARE SQL QUERY
        sql_string = "SELECT * FROM tblitems"

    #DECLARE CONNECTIONS
        oledb_command.Connection = oledb_connection

    #DECLARE QUERY  COMMAND FROM QUERY STRING
        oledb_command.CommandText = sql_string

    #ADAPTER INSERT COMMAND 
        oledb_data_adapter.SelectCommand = oledb_command

    #POPULATING  TO DATA TABLE from DATABASE  using ADAPTER 
        oledb_data_adapter.Fill(data_table)

    #DATA TABLE CONTAINS DATA FROM DB  WILL INSERT TO DATA GRID VIEW 
        dg_view_item.DataSource = data_table
    
    #CLOSING THE ODLEDB CONNECTION AFTER COMMAND
        oledb_connection.Close()

#############################################################################################################################


###################################### SAVE DATA TO MS ACCESS ######################################################


 #open the oledbconnection from MS ACCESS
        oledb_connection.Open()
    #DECLARE SQL QUERY
        sql_string = "INSERT INTO TABLENAME(DB_COLUMN_NAME,DB_COLUMN_NAME2)VALUES('" & TEXTBOX DATA & "','" & TEXTBOX DATA & "');"

    #DECLARE CONNECTIONS
        oledb_command.Connection = oledb_connection
    #DECLARE QUERY  COMMAND FROM QUERY STRING
        oledb_command.CommandText = sql_string
    #EXECUTE THE QUERY, IF  RETURN VALUE  NON 0 , IT WAS SAVED SUCCESSFULLY
        Dim i As Integer  = oledb_command.ExecuteNonQuery
    #CONDITIONAL STATEMENT IF QUERY SUCCESS
        IF I > 0  MsgBox("Item saved!")


################################################################################


###################################### UPDATE DATA TO MS ACCESS ######################################################


 #open the oledbconnection from MS ACCESS
        oledb_connection.Open()
    #DECLARE SQL QUERY
        sql_string = "UPDATE tablename SET DB_COLUMN_NAME1='" & TEXTBOX_DATA1 & "', DB_COLUMN_NAME2='" & TEXTBOX_DATA2 & '" WHERE DB_ID =" & Me.Text(ID Container) & "  "

    #DECLARE CONNECTIONS
        oledb_command.Connection = oledb_connection
    #DECLARE QUERY  COMMAND FROM QUERY STRING
        oledb_command.CommandText = sql_string
    #EXECUTE THE QUERY, IF  RETURN VALUE  NON 0 , IT WAS SAVED SUCCESSFULLY
        Dim i As Integer  = oledb_command.ExecuteNonQuery
    #CONDITIONAL STATEMENT IF QUERY SUCCESS
        IF I > 0  MsgBox("Item saved!")

################################################################################


###################################### UPDATE DATA TO MS ACCESS ######################################################


 #open the oledbconnection from MS ACCESS
        oledb_connection.Open()
    #DECLARE SQL QUERY
        sql_string = "DELETE * FROM db_tblnamme WHERE db_id=" & Me.Text(ID CONTAINER)  & ""

    #DECLARE CONNECTIONS
        oledb_command.Connection = oledb_connection
    #DECLARE QUERY  COMMAND FROM QUERY STRING
        oledb_command.CommandText = sql_string
    #EXECUTE THE QUERY, IF  RETURN VALUE  NON 0 , IT WAS SAVED SUCCESSFULLY
        Dim i As Integer  = oledb_command.ExecuteNonQuery
    #CONDITIONAL STATEMENT IF QUERY SUCCESS
        IF I > 0  MsgBox("Item saved!")

################################################################################









############################################## UTILITY ######################################################################
    
    
        #clearing datagrid view data for refresh realtime:
             dg_view_item.DataSource=Nothing

        #refresh datagrid view from UI
            dg_view_item.refresh()

        #clear data table to insert new updated data from DB
            data_table.Clear()


            #Data gridview must be:
                    AllowUserToAddRows = False
                    AllowUserToResizeRows = False
                    EditMode = EditProgrammatically
                    MultiSelect = False
                    SelectionMode = FullRowSelect
                    Size = New System.Drawing.Size(696, 281)
                    TabIndex = 0
                    #Event Handler:(User clicks)
                        CellContentClick
                        CellClick

            #get data from selected row from DATA GRIDVIEW
                #get current row number 
                    dg_view_item.CurrentRow.Cells(0).Value
                #get text from current ROW selected COLUMN
                    TEXTBOX CONTAINER OR VARIABLE = dg_view_item.CurrentRow.Cells(#CELLNUMBER).Value
            

            #check if OLDEB CONNECTION
                 IF oledb_connection.State = ConnectionState.Open 


            #Always use TRY CATCH in  OLEDB  

#############################################################################################################################










        
