Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class book

        '*********************************************************************
        '  Class Name: book
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

        Private _strBook As String = ""

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

        'Properties For Book
        Public Property Book() As String

            Get

                Return _strBook

            End Get

            Set(ByVal Value As String)

                _strBook = Value

            End Set

        End Property

        Public Sub setBook(ByVal Value As String)

            _strBook = Value

        End Sub
        Public Function getBook() As String

            Return _strBook

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbBook" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetBookByID " & intId)

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

                    _strBook = objCached.getBook

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

                        _strBook = reader(1)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbBook" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteBook " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strBook As String)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertBook '" & strBook & "'"))

                _intId = intId
                _strBook = strBook

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertBook '" & _strBook & "'"))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal strBook As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateBook " & intId & ",'" & strBook & "'")

                _intId = intId
                _strBook = strBook

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateBook " & _intId & ",'" & _strBook & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.book

            Dim strCacheVar As String = "objCachedDbBook" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.book

                objCached.setId(_intId)

                objCached.setBook(_strBook)

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
