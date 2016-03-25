Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class rating

        '*********************************************************************
        '  Class Name: rating
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:50 PM
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

        Private _intPageId As Integer = 0

        Private _intVal As Integer = 0

        Private _strIpaddress As String = ""

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

        'Properties For PageId
        Public Property PageId() As Integer

            Get

                Return _intPageId

            End Get

            Set(ByVal Value As Integer)

                _intPageId = Value

            End Set

        End Property

        Public Sub setPageId(ByVal Value As Integer)

            _intPageId = Value

        End Sub
        Public Function getPageId() As Integer

            Return _intPageId

        End Function

        'Properties For Val
        Public Property Val() As Integer

            Get

                Return _intVal

            End Get

            Set(ByVal Value As Integer)

                _intVal = Value

            End Set

        End Property

        Public Sub setVal(ByVal Value As Integer)

            _intVal = Value

        End Sub
        Public Function getVal() As Integer

            Return _intVal

        End Function

        'Properties For Ipaddress
        Public Property Ipaddress() As String

            Get

                Return _strIpaddress

            End Get

            Set(ByVal Value As String)

                _strIpaddress = Value

            End Set

        End Property

        Public Sub setIpaddress(ByVal Value As String)

            _strIpaddress = Value

        End Sub
        Public Function getIpaddress() As String

            Return _strIpaddress

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbRating" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetRatingByID " & intId)

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

                    _intPageId = objCached.getPageId

                    _intVal = objCached.getVal

                    _strIpaddress = objCached.getIpaddress

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

                        _intPageId = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intVal = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strIpaddress = reader(3)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbRating" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteRating " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intPageId As Integer, ByVal intVal As Integer, ByVal strIpaddress As String)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertRating " & intPageId & "," & intVal & ",'" & strIpaddress & "'"))

                _intId = intId
                _intPageId = intPageId
                _intVal = intVal
                _strIpaddress = strIpaddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertRating " & _intPageId & "," & _intVal & ",'" & _strIpaddress & "'"))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal intPageId As Integer, ByVal intVal As Integer, ByVal strIpaddress As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateRating " & intId & "," & intPageId & "," & intVal & ",'" & strIpaddress & "'")

                _intId = intId
                _intPageId = intPageId
                _intVal = intVal
                _strIpaddress = strIpaddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateRating " & _intId & "," & _intPageId & "," & _intVal & ",'" & _strIpaddress & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.rating

            Dim strCacheVar As String = "objCachedDbRating" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.rating

                objCached.setId(_intId)

                objCached.setPageId(_intPageId)

                objCached.setVal(_intVal)

                objCached.setIpaddress(_strIpaddress)

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
