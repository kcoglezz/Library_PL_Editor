Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class bloqued

        '*********************************************************************
        '  Class Name: bloqued
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intRecordId As Integer)


            If intRecordId <> 0 Then

                getByID(intRecordId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intRecordId As Integer = 0

        Private _strIpAddress As String = ""

#End Region

#Region "Properties"

        'Properties For RecordId
        Public Property RecordId() As Integer

            Get

                Return _intRecordId

            End Get

            Set(ByVal Value As Integer)

                _intRecordId = Value

            End Set

        End Property

        Public Sub setRecordId(ByVal Value As Integer)

            _intRecordId = Value

        End Sub
        Public Function getRecordId() As Integer

            Return _intRecordId

        End Function

        'Properties For IpAddress
        Public Property IpAddress() As String

            Get

                Return _strIpAddress

            End Get

            Set(ByVal Value As String)

                _strIpAddress = Value

            End Set

        End Property

        Public Sub setIpAddress(ByVal Value As String)

            _strIpAddress = Value

        End Sub
        Public Function getIpAddress() As String

            Return _strIpAddress

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intRecordId As Integer)

            Dim strCacheVar As String = "objCachedDbBloqued" & intRecordId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetBloquedByID " & intRecordId)

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

                    _intRecordId = objCached.getRecordId

                    _strIpAddress = objCached.getIpAddress

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

                        _intRecordId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strIpAddress = reader(1)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intRecordId As Integer)

            Dim strCacheVar As String = "objCachedDbBloqued" & intRecordId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteBloqued " & intRecordId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strIpAddress As String)

            Dim intRecordId As Integer = 0

            Try

                intRecordId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertBloqued '" & strIpAddress & "'"))

                _intRecordId = intRecordId
                _strIpAddress = strIpAddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intRecordId As Integer = 0

            Try

                intRecordId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertBloqued '" & _strIpAddress & "'"))

                _intRecordId = intRecordId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intRecordId As Integer, ByVal strIpAddress As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateBloqued " & intRecordId & ",'" & strIpAddress & "'")

                _intRecordId = intRecordId
                _strIpAddress = strIpAddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateBloqued " & _intRecordId & ",'" & _strIpAddress & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.bloqued

            Dim strCacheVar As String = "objCachedDbBloqued" & _intRecordId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.bloqued

                objCached.setRecordId(_intRecordId)

                objCached.setIpAddress(_strIpAddress)

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
