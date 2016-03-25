Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer

    Public Class page_visit

        '*********************************************************************
        '  Class Name: page_visit
        '  Author: Automated Code Generator
        '  Date Created: 4/13/2009 9:38:47 AM
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

        Private _intPageId As Integer = 0

        Private _strIpAddress As String = ""

        Private _dtVisitedOn As Date = #1/1/1900#

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

        'Properties For VisitedOn
        Public Property VisitedOn() As Date

            Get

                Return _dtVisitedOn

            End Get

            Set(ByVal Value As Date)

                _dtVisitedOn = Value

            End Set

        End Property

        Public Sub setVisitedOn(ByVal Value As Date)

            _dtVisitedOn = Value

        End Sub
        Public Function getVisitedOn() As Date

            Return _dtVisitedOn

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intRecordId As Integer)

            Dim strCacheVar As String = "objCachedDbPageVisit" & intRecordId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetPageVisitByID " & intRecordId)

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

                    _intPageId = objCached.getPageId

                    _strIpAddress = objCached.getIpAddress

                    _dtVisitedOn = objCached.getVisitedOn

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

                        _intPageId = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _strIpAddress = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _dtVisitedOn = reader(3)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intRecordId As Integer)

            Dim strCacheVar As String = "objCachedDbPageVisit" & intRecordId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeletePageVisit " & intRecordId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intPageId As Integer, ByVal strIpAddress As String, ByVal dtVisitedOn As Date)

            Dim intRecordId As Integer = 0

            Try

                intRecordId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertPageVisit " & intPageId & ",'" & strIpAddress & "','" & dtVisitedOn & "'"))

                _intRecordId = intRecordId
                _intPageId = intPageId
                _strIpAddress = strIpAddress
                _dtVisitedOn = dtVisitedOn

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intRecordId As Integer = 0

            Try

                intRecordId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertPageVisit " & _intPageId & ",'" & _strIpAddress & "','" & _dtVisitedOn & "'"))

                _intRecordId = intRecordId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intRecordId As Integer, ByVal intPageId As Integer, ByVal strIpAddress As String, ByVal dtVisitedOn As Date)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdatePageVisit " & intRecordId & "," & intPageId & ",'" & strIpAddress & "','" & dtVisitedOn & "'")

                _intRecordId = intRecordId
                _intPageId = intPageId
                _strIpAddress = strIpAddress
                _dtVisitedOn = dtVisitedOn

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdatePageVisit " & _intRecordId & "," & _intPageId & ",'" & _strIpAddress & "','" & _dtVisitedOn & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.page_visit

            Dim strCacheVar As String = "objCachedDbPageVisit" & _intRecordId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.page_visit

                objCached.setRecordId(_intRecordId)

                objCached.setPageId(_intPageId)

                objCached.setIpAddress(_strIpAddress)

                objCached.setVisitedOn(_dtVisitedOn)

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
