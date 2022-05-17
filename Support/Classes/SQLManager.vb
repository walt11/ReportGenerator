Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Configuration.ConfigurationSettings

Public Class SQLManager

    'https://www.dreamincode.net/forums/topic/32392-sql-basics-in-vbnet/

    ''' <summary>
    ''' Function to retrieve the connection from the app.config
    ''' </summary>
    ''' <param name="conName">Name of the connectionString to retrieve</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetConnectionString(ByVal conName As String) As String
        'variable to hold our connection string for returning it
        Dim strReturn As New String("")
        'check to see if the user provided a connection string name
        'this is for if your application has more than one connection string
        If Not String.IsNullOrEmpty(conName) Then
            'a connection string name was provided
            'get the connection string by the name provided
            strReturn = ConfigurationManager.ConnectionStrings(conName).ConnectionString
        Else
            'no connection string name was provided
            'get the default connection string
            strReturn = ConfigurationManager.ConnectionStrings("YourConnectionName").ConnectionString
        End If
        'return the connection string to the calling method
        Return strReturn
    End Function

    ''' <summary>
    ''' Returns a BindingSource, which is used with, for example, a DataGridView control
    ''' </summary>
    ''' <param name="cmd">"pre-Loaded" command, ready to be executed</param>
    ''' <returns>BindingSource</returns>
    ''' <remarks>Use this function to ease populating controls that use a BindingSource</remarks>
    Public Shared Function GetBindingSource(ByVal cmd As SqlCommand) As BindingSource
        'declare our binding source
        Dim oBindingSource As New BindingSource()
        ' Create a new data adapter based on the specified query.
        Dim daGet As New SqlDataAdapter(cmd)
        ' Populate a new data table and bind it to the BindingSource.
        Dim dtGet As New DataTable()
        'set the timeout of the SqlCommandObject
        cmd.CommandTimeout = 240
        dtGet.Locale = System.Globalization.CultureInfo.InvariantCulture
        Try
            'fill the DataTable with the SqlDataAdapter
            daGet.Fill(dtGet)
        Catch ex As Exception
            'check for errors
            MsgBox(ex.Message, "Error in GetBindingSource")
            Return Nothing
        End Try
        'set the DataSource for the BindingSource to the DataTable
        oBindingSource.DataSource = dtGet
        'return the BindingSource to the calling method or control
        Return oBindingSource
    End Function

    ''' <summary>
    ''' Method for handling the ConnectionState of 
    ''' the connection object passed to it
    ''' </summary>
    ''' <param name="conn">The SqlConnection Object</param>
    ''' <remarks></remarks>
    Public Shared Sub HandleConnection(ByVal conn As SqlConnection)
        With conn
            'do a switch on the state of the connection
            Select Case .State
                Case ConnectionState.Open
                    'the connection is open
                    'close then re-open
                    .Close()
                    .Open()
                    Exit Select
                Case ConnectionState.Closed
                    'connection is open
                    'open the connection
                    .Open()
                    Exit Select
                Case Else
                    .Close()
                    .Open()
                    Exit Select
            End Select
        End With
    End Sub

    Public Shared Function InsertNewRecord(ByVal item1 As String, ByVal item2 As String, ByVal item3 As String) As Boolean
        'Create the objects we need to insert a new record
        Dim cnInsert As New SqlConnection(GetConnectionString("YourConnName"))
        Dim cmdInsert As New SqlCommand
        Dim sSQL As New String("")
        Dim iSqlStatus As Integer

        'Set the stored procedure we're going to execute
        sSQL = "YourProcName"

        'Inline sql needs to be structured like so
        'sSQL = "INSERT INTO YourTable(column1,column2,column3) VALUES('" & item1 & "','" & item2 & "','" & item3 & "')"

        'Clear any parameters
        cmdInsert.Parameters.Clear()
        Try
            'Set the SqlCommand Object Properties
            With cmdInsert
                'Tell it what to execute
                .CommandText = sSQL 'Your sql statement if using inline sql
                'Tell it its a stored procedure
                .CommandType = CommandType.StoredProcedure 'CommandType.Text for inline sql
                'If you are indeed using a stored procedure
                'the next 3 lines pertain to you
                'Now add the parameters to our procedure
                'NOTE: Replace @value1.... with your parameter names in your stored procedure
                'and add all your parameters in this fashion
                .Parameters.AddWithValue("@value1", item1)
                .Parameters.AddWithValue("@value2", item2)
                .Parameters.AddWithValue("@value3", item3)
                'Set the connection of the object
                .Connection = cnInsert
            End With

            'Now take care of the connection
            HandleConnection(cnInsert)

            'Set the iSqlStatus to the ExecuteNonQuery status of the insert (0 = success, 1 = failed)
            iSqlStatus = cmdInsert.ExecuteNonQuery

            'Now check the status
            If Not iSqlStatus = 0 Then
                'DO your failed messaging here
                Return False
            Else
                'Do your success work here
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, "Error")
        Finally
            'Now close the connection
            HandleConnection(cnInsert)
        End Try

    End Function

    Public Shared Function DeleteRecord(ByVal id As Integer) As Boolean
        'Create the objects we need to insert a new record
        Dim cnDelete As New SqlConnection(GetConnectionString("YourConnName"))
        Dim cmdDelete As New SqlCommand
        Dim sSQL As New String("")
        Dim iSqlStatus As Integer

        'Set the stored procedure we're going to execute
        sSQL = "YourProcName"

        'Inline sql needs to be structured like so
        'sSQL = "DELETE FROM YourTable WHERE YourID = " & id

        'Clear any parameters
        cmdDelete.Parameters.Clear()
        Try
            'Set the SqlCommand Object Properties
            With cmdDelete
                'Tell it what to execute
                .CommandText = sSQL 'Your sql statement if using inline sql
                'Tell it its a stored procedure
                .CommandType = CommandType.StoredProcedure 'CommandType.Text for inline sql
                'If you are indeed using a stored procedure
                'the next 3 lines pertain to you
                'Now add the parameters to our procedure
                'NOTE: Replace @value1.... with your parameter names in your stored procedure
                'and add all your parameters in this fashion
                .Parameters.AddWithValue("@YourID", id)
                'Set the connection of the object
                .Connection = cnDelete
            End With

            'Now take care of the connection
            HandleConnection(cnDelete)

            'Set the iSqlStatus to the ExecuteNonQuery 
            'status of the insert (0 = success, 1 = failed)
            iSqlStatus = cmdDelete.ExecuteNonQuery

            'Now check the status
            If Not iSqlStatus = 0 Then
                'DO your failed messaging here
                Return False
            Else
                'Do your success work here
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, "Error")
            Return False
        Finally
            'Now close the connection
            HandleConnection(cnDelete)
        End Try

    End Function

    Public Shared Function UpdateRecord(ByVal item1 As String, ByVal item2 As String, ByVal id As Integer) As Boolean
        'Create the objects we need to insert a new record
        Dim cnUpdate As New SqlConnection(GetConnectionString("YourConnName"))
        Dim cmdUpdate As New SqlCommand
        Dim sSQL As New String("")
        Dim iSqlStatus As Integer

        'Set the stored procedure we're going to execute
        sSQL = "YourProcName"

        'Inline sql needs to be structured like so
        'sSQL = "UPDATE YourTable SET column1 = '" & item1 & "',column2 = '" & item2 & "' WHERE YourId = " & id

        'Clear any parameters
        cmdUpdate.Parameters.Clear()
        Try
            'Set the SqlCommand Object Properties
            With cmdUpdate
                'Tell it what to execute
                .CommandText = sSQL 'Your sql statement if using inline sql
                'Tell it its a stored procedure
                .CommandType = CommandType.StoredProcedure 'CommandType.Text for inline sql
                'If you are indeed using a stored procedure
                'the next 3 lines pertain to you
                'Now add the parameters to our procedure
                'NOTE: Replace @value1.... with your parameter names in your stored procedure
                'and add all your parameters in this fashion
                .Parameters.AddWithValue("@value1", item1)
                .Parameters.AddWithValue("@value2", item2)
                .Parameters.AddWithValue("@YourID", id)
                'Set the connection of the object
                .Connection = cnUpdate
            End With

            'Now take care of the connection
            HandleConnection(cnUpdate)

            'Set the iSqlStatus to the ExecuteNonQuery 
            'status of the insert (0 = success, 1 = failed)
            iSqlStatus = cmdUpdate.ExecuteNonQuery

            'Now check the status
            If Not iSqlStatus = 0 Then
                'DO your failed messaging here
                Return False
            Else
                'Do your success work here
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, "Error")
        Finally
            'Now close the connection
            HandleConnection(cnUpdate)
        End Try
    End Function

    Public Shared Function GetRecordsByID(ByVal value As Integer) As BindingSource
        'The value that will be passed to the Command Object (this is a stored procedure)
        Dim sSQL As String = "YourProcName"
        'If using inline sql format is as such
        'sSQL = "SELECT value1,value2,value3 FROM YourTable WHERE YourValue = " & value
        'Stored procedure to execute
        Dim cnGetRecords As New SqlConnection(GetConnectionString("YourConnectionName"))
        'SqlConnection Object to use
        Dim cmdGetRecords As New SqlCommand()
        'SqlCommand Object to use
        Dim daGetRecords As New SqlDataAdapter()
        Dim dsGetRecords As New DataSet()
        'Clear any parameters
        cmdGetRecords.Parameters.Clear()
        Try
            With cmdGetRecords
                'set the SqlCommand Object Parameters
                .CommandText = sSQL
                'tell it what to execute
                .CommandType = CommandType.StoredProcedure
                'tell it its executing a Stored Procedure
                'heres the difference from the last method
                'here we are adding a parameter to send to our stored procedure
                'you use the AddWithValue, then the name of the parameter in your stored procedure
                'then the variable that holds that value
                .Parameters.AddWithValue("@year", value)
                'Set the Connection for the Command Object
                .Connection = cnGetRecords
            End With
            'set the state of the SqlConnection Object
            HandleConnection(cnGetRecords)
            'create BindingSource to return for our DataGrid Control
            Dim oBindingSource As BindingSource = GetBindingSource(cmdGetRecords)
            'now check to make sure a BindingSource was returned
            If Not oBindingSource Is Nothing Then
                'return the binding source to the calling method
                Return oBindingSource
            Else
                'no binding source was returned
                'let the user know the error
                Throw New Exception("There was no BindingSource returned")
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, "Error Retrieving Data")
            Return Nothing
        Finally
            HandleConnection(cnGetRecords)
        End Try
    End Function

    Public Shared Function GetRecords() As BindingSource
        'The value that will be passed to the Command Object (this is a stored procedure)
        Dim sSQL As String = "YourProcName"
        'If using inline sql format is as such
        'sSQL = "SELECT * FROM YourTable
        'Stored procedure to execute
        Dim cnGetRecords As New SqlConnection(GetConnectionString("YourConnectionName"))
        'SqlConnection Object to use
        Dim cmdGetRecords As New SqlCommand()
        'SqlCommand Object to use
        Dim daGetRecords As New SqlDataAdapter()
        Dim dsGetRecords As New DataSet()
        'Clear any parameters
        cmdGetRecords.Parameters.Clear()
        Try
            With cmdGetRecords
                'set the SqlCommand Object Parameters
                .CommandText = sSQL
                'tell it what to execute
                .CommandType = CommandType.StoredProcedure
                'Set the Connection for the Command Object
                .Connection = cnGetRecords
            End With
            'set the state of the SqlConnection Object
            HandleConnection(cnGetRecords)
            'create BindingSource to return for our DataGrid Control
            Dim oBindingSource As BindingSource = GetBindingSource(cmdGetRecords)
            'now check to make sure a BindingSource was returned
            If Not oBindingSource Is Nothing Then
                'return the binding source to the calling method
                Return oBindingSource
            Else
                'no binding source was returned
                'let the user know the error
                Throw New Exception("There was no BindingSource returned")
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, "Error Retrieving Data")
            Return Nothing
        Finally
            HandleConnection(cnGetRecords)
        End Try
    End Function
End Class
