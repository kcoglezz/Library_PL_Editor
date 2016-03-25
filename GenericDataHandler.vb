Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Xml
Imports System.Text
Imports System.Web
Imports Sistema.PL.Entidad

Namespace ParaLideres

    Public Class GenericDataHandler

#Region "Global Variables"

        Private Shared mContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Private Shared oConex As New InfoConexion()
        Private Shared strConnection As String = oConex.StringConnection
        Private Shared _develop As String = System.Web.Configuration.WebConfigurationManager.AppSettings("DeveloperAccount")

#End Region

#Region "Subs"

        Sub New()


        End Sub

        Protected Overrides Sub Finalize()

            mContext = Nothing

        End Sub

#End Region

#Region "Properties"

        'Public Shared Property ServerName() As String

        '    Get

        '        Return strServer

        '    End Get

        '    Set(ByVal Value As String)

        '        strServer = Value

        '    End Set

        'End Property

        Public Shared ReadOnly Property ConnectionString() As String

            Get

                Return strConnection

            End Get

        End Property

        'Public Shared Property DatabaseName() As String

        '    Get

        '        Return strDatabase

        '    End Get

        '    Set(ByVal Value As String)

        '        strDatabase = Value

        '    End Set

        'End Property

        'Public Shared Property Username() As String

        '    Get

        '        Return strUsername

        '    End Get

        '    Set(ByVal Value As String)

        '        strUsername = Value

        '    End Set

        'End Property

        'Public Shared Property Password() As String
        '    Get

        '        Return strPassword

        '    End Get

        '    Set(ByVal Value As String)

        '        strPassword = Value

        '    End Set
        'End Property

        'Public Shared Property MaxPoolSize() As String
        '    Get

        '        Return strMaxPoolSize

        '    End Get

        '    Set(ByVal Value As String)

        '        strMaxPoolSize = Value

        '    End Set
        'End Property

#End Region

#Region "Functions"

        Public Shared Function BindData(ByVal sqlString As String, ByVal controlObj As Object) As Integer

            Dim mConn As New SqlConnection
            Dim mDR As SqlDataReader
            Dim cmd As SqlCommand
            Dim records As Integer = 0


            Try

                mConn = New SqlConnection(strConnection)
                mConn.Open()
                cmd = New SqlCommand(sqlString, mConn)

                mDR = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                controlObj.DataSource = mDR

                Select Case controlObj.GetType.Name

                    Case "DataGrid", "DataList", "Repeater", "SimpleDataGrid"

                        'do nothing

                    Case Else

                        controlObj.DataValueField = mDR.GetName(0)
                        controlObj.DataTextField = mDR.GetName(1)

                End Select

                controlObj.DataBind()
                records = controlObj.Items.Count()

                Return records

            Catch ex As Exception

                mContext.Trace.Write("Generic Data Handler Error: " & ex.Source & " " & ex.Message & " " & ex.ToString() & " " & ex.InnerException.ToString())
                Return -1

            Finally

                mDR.Close()
                cmd = Nothing
                mConn = Nothing
                mDR = Nothing

            End Try

        End Function

        Public Shared Function ExecSQL(ByVal mySQL As String) As Integer

            'after executing query returns number of rows affected

            Dim mConn As New SqlConnection(strConnection)
            Dim sqlCmd As SqlCommand

            Try

                mConn.Open()
                sqlCmd = New SqlCommand(mySQL, mConn)

                Return sqlCmd.ExecuteNonQuery()

            Catch ex As Exception

                Functions.ShowError(ex, "ExecSQL")

                Return -1

            Finally

                mConn.Close()
                mConn = Nothing
                sqlCmd = Nothing

            End Try

        End Function

        Public Shared Function ExecSQL(ByVal cmdStoredProc As SqlCommand) As Integer

            'after executing query returns number of rows affected
            Dim mConn As New SqlConnection(strConnection)
            Dim cmd As SqlCommand

            Try

                mConn.Open()

                cmd = cmdStoredProc

                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecSQL: " & cmd.CommandText)
                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecSQL: " & cmd.CommandType.ToString())

                cmd.Connection = mConn

                Return cmd.ExecuteNonQuery

            Catch ex As Exception

                Functions.ShowError(ex, "ExecSQL")

            Finally

                mConn.Close()
                cmd = Nothing
                mConn = Nothing

            End Try

        End Function

        Public Shared Function ExecScalar(ByVal mySQL As String) As Object

            'returns only the first column of first row, the rest columns and rows are ignored

            Dim mConn As New SqlConnection(strConnection)
            Dim sqlCmd As SqlCommand

            Try

                mConn.Open()

                sqlCmd = New SqlCommand(mySQL, mConn)

                Return sqlCmd.ExecuteScalar()

            Catch ex As Exception

                Functions.SendMail("apoyo@ParaLideres.org", _develop, "Error on PL ExecScalar", "Error while trying to execute " & mySQL)

                Functions.ShowError(ex, "ExecScalar")

            Finally

                mConn.Close()
                mConn = Nothing
                sqlCmd = Nothing

            End Try

        End Function

        Public Shared Function ExecScalar(ByVal cmdStoredProc As SqlCommand) As Object

            Dim mConn As New SqlConnection(strConnection)
            Dim cmd As SqlCommand
            Dim objReturn As Object


            Try

                mConn.Open()

                cmd = cmdStoredProc

                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecScalar: " & cmd.CommandText)
                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecScalar: " & cmd.CommandType.ToString())

                cmd.Connection = mConn

                objReturn = cmd.ExecuteScalar()

                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecScalar-> Obj type: " & objReturn.GetType.ToString)

                System.Web.HttpContext.Current.Trace.Write("GDH -> ExecScalar-> Obj: " & objReturn.ToString())

                Return objReturn

            Catch ex As Exception

                Functions.ShowError(ex, "ExecScalar")

            Finally

                mConn.Close()
                mConn = Nothing
                cmd = Nothing

            End Try

        End Function

        Public Shared Function GetRecords(ByVal sqlString As String, Optional ByVal intCommandSecondsTimeout As Integer = 240) As SqlDataReader

            Dim mConn As New SqlConnection(strConnection)
            Dim cmd As SqlCommand

            Try

                mConn.Open()

                cmd = New SqlCommand(sqlString, mConn)
                cmd.CommandTimeout = intCommandSecondsTimeout

                Return cmd.ExecuteReader(CommandBehavior.CloseConnection)

            Catch ex As Exception

                Functions.ShowError(ex, "GetRecords")

            Finally

                cmd = Nothing
                mConn = Nothing

            End Try

        End Function

        Public Shared Function GetRecords(ByVal cmdStoredProc As SqlCommand, Optional ByVal intCommandSecondsTimeout As Integer = 240) As SqlDataReader

            Dim mConn As New SqlConnection(strConnection)
            Dim cmd As SqlCommand

            Try

                mConn.Open()

                cmd = cmdStoredProc
                cmd.CommandTimeout = intCommandSecondsTimeout
                cmd.Connection = mConn

                Return cmd.ExecuteReader(CommandBehavior.CloseConnection)

            Catch ex As Exception

                Functions.ShowError(ex, "GetRecords")

            Finally

                cmd = Nothing
                mConn = Nothing

            End Try

        End Function

        Public Shared Function GetDataSet(ByVal sqlString As String, Optional ByVal mTableName As String = "mTable", Optional ByVal mStartRecord As Integer = 0, Optional ByVal mMaxRecords As Integer = 0) As DataSet

            Dim mConn As New SqlConnection
            mConn.ConnectionString = strConnection

            Try
                mConn.Open()
            Catch e As SqlException

                mContext.Response.Write("<p align=center><font color=red size=4>Error Found!</font></p><p>The application was not able to connect to the database: " & System.Web.Configuration.WebConfigurationManager.AppSettings("Database") & " on server: " & System.Web.Configuration.WebConfigurationManager.AppSettings("Server"))
                mContext.Response.Write("<p>Information for support personnel<br> Error generated by Generic Data Handler on SQL Server: " & e.Message)
                mContext.Response.Write("<p>Please check the settings of the config.web file to make sure that you are pointing to the right database and server and make sure that you can connect to the specified server")

            End Try

            Dim mSqlAdapter As SqlDataAdapter = New SqlDataAdapter(sqlString, mConn)
            Dim mDataSet As DataSet = New DataSet

            Try

                If mMaxRecords <> 0 Then
                    mSqlAdapter.Fill(mDataSet, mStartRecord, mMaxRecords, mTableName)
                Else
                    mSqlAdapter.Fill(mDataSet, mTableName)
                End If

                Return mDataSet

            Catch e As SqlException

                mContext.Response.Write("Generic Data Handler Error (SQL Server): " & e.Message)

            Finally

                mConn = Nothing
                mSqlAdapter = Nothing
                mDataSet = Nothing

            End Try

        End Function

        Public Shared Function GetTableSchema(ByVal strSQLQuery As String) As String

            Dim cn As New SqlConnection
            Dim cmd As New SqlCommand
            Dim schemaTable As DataTable
            Dim myReader As SqlDataReader
            Dim myField As DataRow
            Dim myProperty As DataColumn
            Dim sb As StringBuilder = New StringBuilder("")

            'Open a connection to the SQL Server 
            cn.ConnectionString = strConnection
            cn.Open()

            'Retrieve records from the Employees table into a DataReader.
            cmd.Connection = cn
            cmd.CommandText = strSQLQuery
            myReader = cmd.ExecuteReader(CommandBehavior.KeyInfo)

            'Retrieve column schema into a DataTable.
            schemaTable = myReader.GetSchemaTable()

            'Xavier logic

            sb.Append("<table border=1 width=500 cellpadding=1 cellspacing=0>")
            sb.Append("<tr class=BOLD>")

            'table headers
            sb.Append("<tr class=BOLD>")
            For Each myProperty In schemaTable.Columns

                sb.Append("<td>" & UCase(myProperty.ColumnName) & "</td>")

            Next
            sb.Append("</tr>")


            'table values
            For Each myField In schemaTable.Rows
                sb.Append("<tr class=GEN>")
                For Each myProperty In schemaTable.Columns
                    sb.Append("<td>" & myField(myProperty).ToString() & "</td>")
                Next
                sb.Append("</tr>")
            Next

            sb.Append("</table>")

            'Always close the DataReader and Connection objects.
            myReader.Close()
            cn.Close()

            cmd = Nothing
            cn = Nothing
            myReader = Nothing

            Return sb.ToString

        End Function

        Public Shared Function TestConnection() As Boolean

            Dim intMinutesInterval As Integer = 5

            Dim blCheck As Boolean = False

            If IsNothing(HttpContext.Current.Cache.Get("LastSQLCheck")) Then

                HttpContext.Current.Trace.Warn("1 was true")

                blCheck = ConnectToDatabase()

            ElseIf CDate(HttpContext.Current.Cache.Get("LastSQLCheck")).AddMinutes(intMinutesInterval) < Date.Now Then

                HttpContext.Current.Trace.Warn("2 was true")

                blCheck = ConnectToDatabase()

            ElseIf IsNothing(HttpContext.Current.Cache.Get("CanConnectToSql")) Then

                HttpContext.Current.Trace.Warn("3 was true")

                blCheck = ConnectToDatabase()

            Else

                HttpContext.Current.Trace.Warn("4 was true")

                Try

                    blCheck = CBool(HttpContext.Current.Cache.Get("CanConnectToSql"))

                Catch ex As Exception

                    blCheck = False

                End Try

            End If

            Return blCheck

        End Function


        Public Shared Function ConnectToDatabase() As Boolean

            Dim mConn As New SqlConnection(strConnection)

            Try

                HttpContext.Current.Trace.Warn("ConnectToDatabase was called!")

                HttpContext.Current.Cache.Insert("LastSQLCheck", Date.Now)

                Try

                    mConn.Open()

                Catch ex As Exception

                    HttpContext.Current.Trace.Warn(ex.ToString())

                End Try

                If mConn.State = ConnectionState.Open Then

                    HttpContext.Current.Cache.Insert("CanConnectToSql", True)

                    mConn.Close()

                    Return True

                Else

                    HttpContext.Current.Cache.Insert("CanConnectToSql", False)

                    Return False

                End If

            Catch ex As Exception

                HttpContext.Current.Trace.Warn(ex.ToString())

            Finally

                mConn = Nothing

            End Try

        End Function

#End Region

    End Class

End Namespace
