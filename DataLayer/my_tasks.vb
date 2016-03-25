Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class my_tasks

        '*********************************************************************
        '  Class Name: my_tasks
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intId As Integer)


            If intId <> 0 Then

                getByID(intId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intId As Integer = 0

        Private _intTaskBy As Integer = 0

        Private _intTaskTo As Integer = 0

        Private _strTitle As String = ""

        Private _strBody As String = ""

        Private _intPriority As Integer = 0

        Private _dtDateSubmited As Date = #1/1/1900#

        Private _dtDueDate As Date = #1/1/1900#

        Private _intStatus As Integer = 0

#End Region

#Region "Properties"

        'Properties For Id
        Public Property Id() As Integer

            Get

                Return _intId

            End Get

            Set(ByVal Value As Integer)

                _intId = Value

            End Set

        End Property

        Public Sub setId(ByVal Value As Integer)

            _intId = Value

        End Sub
        Public Function getId() As Integer

            Return _intId

        End Function

        'Properties For TaskBy
        Public Property TaskBy() As Integer

            Get

                Return _intTaskBy

            End Get

            Set(ByVal Value As Integer)

                _intTaskBy = Value

            End Set

        End Property

        Public Sub setTaskBy(ByVal Value As Integer)

            _intTaskBy = Value

        End Sub
        Public Function getTaskBy() As Integer

            Return _intTaskBy

        End Function

        'Properties For TaskTo
        Public Property TaskTo() As Integer

            Get

                Return _intTaskTo

            End Get

            Set(ByVal Value As Integer)

                _intTaskTo = Value

            End Set

        End Property

        Public Sub setTaskTo(ByVal Value As Integer)

            _intTaskTo = Value

        End Sub
        Public Function getTaskTo() As Integer

            Return _intTaskTo

        End Function

        'Properties For Title
        Public Property Title() As String

            Get

                Return _strTitle

            End Get

            Set(ByVal Value As String)

                _strTitle = Value

            End Set

        End Property

        Public Sub setTitle(ByVal Value As String)

            _strTitle = Value

        End Sub
        Public Function getTitle() As String

            Return _strTitle

        End Function

        'Properties For Body
        Public Property Body() As String

            Get

                Return _strBody

            End Get

            Set(ByVal Value As String)

                _strBody = Value

            End Set

        End Property

        Public Sub setBody(ByVal Value As String)

            _strBody = Value

        End Sub
        Public Function getBody() As String

            Return _strBody

        End Function

        'Properties For Priority
        Public Property Priority() As Integer

            Get

                Return _intPriority

            End Get

            Set(ByVal Value As Integer)

                _intPriority = Value

            End Set

        End Property

        Public Sub setPriority(ByVal Value As Integer)

            _intPriority = Value

        End Sub
        Public Function getPriority() As Integer

            Return _intPriority

        End Function

        'Properties For DateSubmited
        Public Property DateSubmited() As Date

            Get

                Return _dtDateSubmited

            End Get

            Set(ByVal Value As Date)

                _dtDateSubmited = Value

            End Set

        End Property

        Public Sub setDateSubmited(ByVal Value As Date)

            _dtDateSubmited = Value

        End Sub
        Public Function getDateSubmited() As Date

            Return _dtDateSubmited

        End Function

        'Properties For DueDate
        Public Property DueDate() As Date

            Get

                Return _dtDueDate

            End Get

            Set(ByVal Value As Date)

                _dtDueDate = Value

            End Set

        End Property

        Public Sub setDueDate(ByVal Value As Date)

            _dtDueDate = Value

        End Sub
        Public Function getDueDate() As Date

            Return _dtDueDate

        End Function

        'Properties For Status
        Public Property Status() As Integer

            Get

                Return _intStatus

            End Get

            Set(ByVal Value As Integer)

                _intStatus = Value

            End Set

        End Property

        Public Sub setStatus(ByVal Value As Integer)

            _intStatus = Value

        End Sub
        Public Function getStatus() As Integer

            Return _intStatus

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbMyTasks" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMyTasksByID " & intId)

                    setValues(reader)

                    HttpContext.Current.Trace.Write("Data Layer Object retrieved database")

                    HttpContext.Current.Cache.Insert(strCacheVar, Me)

                Catch ex As Exception

                    Throw ex

                Finally

                    reader.Close()
                    reader = Nothing
                End Try

            Else 'If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim objCached As Object = HttpContext.Current.Cache.Get(strCacheVar)

                Try

                    HttpContext.Current.Trace.Write("Data Layer Object retrieved from cache")

                    _intId = objCached.getId

                    _intTaskBy = objCached.getTaskBy

                    _intTaskTo = objCached.getTaskTo

                    _strTitle = objCached.getTitle

                    _strBody = objCached.getBody

                    _intPriority = objCached.getPriority

                    _dtDateSubmited = objCached.getDateSubmited

                    _dtDueDate = objCached.getDueDate

                    _intStatus = objCached.getStatus

                Catch ex As Exception

                    Throw ex

                Finally

                    objCached = Nothing

                End Try

            End If 'If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

        End Sub


        'Sub setValues:  Sets the values to all properties for this object from database

        Private Sub setValues(ByVal reader As System.Data.SqlClient.SqlDataReader)
            Try
                If reader.HasRows() Then

                    reader.Read()

                    If Not reader.IsDBNull(0) Then

                        _intId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _intTaskBy = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intTaskTo = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strTitle = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strBody = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _intPriority = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _dtDateSubmited = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _dtDueDate = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _intStatus = reader(8)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbMyTasks" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteMyTasks " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intTaskBy As Integer, ByVal intTaskTo As Integer, ByVal strTitle As String, ByVal strBody As String, ByVal intPriority As Integer, ByVal dtDateSubmited As Date, ByVal dtDueDate As Date, ByVal intStatus As Integer)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertMyTasks " & intTaskBy & "," & intTaskTo & ",'" & strTitle & "','" & strBody & "'," & intPriority & ",'" & dtDateSubmited & "','" & dtDueDate & "'," & intStatus))

                _intId = intId
                _intTaskBy = intTaskBy
                _intTaskTo = intTaskTo
                _strTitle = strTitle
                _strBody = strBody
                _intPriority = intPriority
                _dtDateSubmited = dtDateSubmited
                _dtDueDate = dtDueDate
                _intStatus = intStatus

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertMyTasks " & _intTaskBy & "," & _intTaskTo & ",'" & _strTitle & "','" & _strBody & "'," & _intPriority & ",'" & _dtDateSubmited & "','" & _dtDueDate & "'," & _intStatus))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal intTaskBy As Integer, ByVal intTaskTo As Integer, ByVal strTitle As String, ByVal strBody As String, ByVal intPriority As Integer, ByVal dtDateSubmited As Date, ByVal dtDueDate As Date, ByVal intStatus As Integer)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateMyTasks " & intId & "," & intTaskBy & "," & intTaskTo & ",'" & strTitle & "','" & strBody & "'," & intPriority & ",'" & dtDateSubmited & "','" & dtDueDate & "'," & intStatus)

                _intId = intId
                _intTaskBy = intTaskBy
                _intTaskTo = intTaskTo
                _strTitle = strTitle
                _strBody = strBody
                _intPriority = intPriority
                _dtDateSubmited = dtDateSubmited
                _dtDueDate = dtDueDate
                _intStatus = intStatus

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateMyTasks " & _intId & "," & _intTaskBy & "," & _intTaskTo & ",'" & _strTitle & "','" & _strBody & "'," & _intPriority & ",'" & _dtDateSubmited & "','" & _dtDueDate & "'," & _intStatus)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.my_tasks

            Dim strCacheVar As String = "objCachedDbMyTasks" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.my_tasks

                objCached.setId(_intId)

                objCached.setTaskBy(_intTaskBy)

                objCached.setTaskTo(_intTaskTo)

                objCached.setTitle(_strTitle)

                objCached.setBody(_strBody)

                objCached.setPriority(_intPriority)

                objCached.setDateSubmited(_dtDateSubmited)

                objCached.setDueDate(_dtDueDate)

                objCached.setStatus(_intStatus)

                HttpContext.Current.Cache.Insert(strCacheVar, objCached)

            Catch ex As Exception

                Throw ex

            Finally

                objCached = Nothing

            End Try

        End Sub




#End Region


    End Class

End Namespace
