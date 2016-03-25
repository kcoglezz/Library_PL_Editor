Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class section_articles

        '*********************************************************************
        '  Class Name: section_articles
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

        Private _intSectionId As Integer = 0

        Private _intArticles As Integer = 0

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

        'Properties For SectionId
        Public Property SectionId() As Integer

            Get

                Return _intSectionId

            End Get

            Set(ByVal Value As Integer)

                _intSectionId = Value

            End Set

        End Property

        Public Sub setSectionId(ByVal Value As Integer)

            _intSectionId = Value

        End Sub
        Public Function getSectionId() As Integer

            Return _intSectionId

        End Function

        'Properties For Articles
        Public Property Articles() As Integer

            Get

                Return _intArticles

            End Get

            Set(ByVal Value As Integer)

                _intArticles = Value

            End Set

        End Property

        Public Sub setArticles(ByVal Value As Integer)

            _intArticles = Value

        End Sub
        Public Function getArticles() As Integer

            Return _intArticles

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbSectionArticles" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionArticlesByID " & intId)

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

                    _intSectionId = objCached.getSectionId

                    _intArticles = objCached.getArticles

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

                        _intSectionId = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intArticles = reader(2)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbSectionArticles" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteSectionArticles " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intSectionId As Integer, ByVal intArticles As Integer)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertSectionArticles " & intSectionId & "," & intArticles))

                _intId = intId
                _intSectionId = intSectionId
                _intArticles = intArticles

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertSectionArticles " & _intSectionId & "," & _intArticles))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal intSectionId As Integer, ByVal intArticles As Integer)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateSectionArticles " & intId & "," & intSectionId & "," & intArticles)

                _intId = intId
                _intSectionId = intSectionId
                _intArticles = intArticles

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateSectionArticles " & _intId & "," & _intSectionId & "," & _intArticles)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.section_articles

            Dim strCacheVar As String = "objCachedDbSectionArticles" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.section_articles

                objCached.setId(_intId)

                objCached.setSectionId(_intSectionId)

                objCached.setArticles(_intArticles)

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
