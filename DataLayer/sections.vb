Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class sections

        '*********************************************************************
        '  Class Name: sections
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:50 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intSectionId As Integer)


            If intSectionId <> 0 Then

                getByID(intSectionId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intSectionId As Integer = 0

        Private _strSectionName As String = ""

        Private _intSectionParent As Integer = 0

        Private _btPostInMenu As Byte = 0

        Private _strBlurb As String = ""

        Private _strPagetext As String = ""

        Private _intUserId As Integer = 0

        Private _dtLastModified As Date = #1/1/1900#

        Private _intModBy As Integer = 0

        Private _boolShowOnDroplist As Boolean = False

        Private _intArticleCount As Integer = 0

#End Region

#Region "Properties"

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

        'Properties For SectionName
        Public Property SectionName() As String

            Get

                Return _strSectionName

            End Get

            Set(ByVal Value As String)

                _strSectionName = Value

            End Set

        End Property

        Public Sub setSectionName(ByVal Value As String)

            _strSectionName = Value

        End Sub
        Public Function getSectionName() As String

            Return _strSectionName

        End Function

        'Properties For SectionParent
        Public Property SectionParent() As Integer

            Get

                Return _intSectionParent

            End Get

            Set(ByVal Value As Integer)

                _intSectionParent = Value

            End Set

        End Property

        Public Sub setSectionParent(ByVal Value As Integer)

            _intSectionParent = Value

        End Sub
        Public Function getSectionParent() As Integer

            Return _intSectionParent

        End Function

        'Properties For PostInMenu
        Public Property PostInMenu() As Byte

            Get

                Return _btPostInMenu

            End Get

            Set(ByVal Value As Byte)

                _btPostInMenu = Value

            End Set

        End Property

        Public Sub setPostInMenu(ByVal Value As Byte)

            _btPostInMenu = Value

        End Sub
        Public Function getPostInMenu() As Byte

            Return _btPostInMenu

        End Function

        'Properties For Blurb
        Public Property Blurb() As String

            Get

                Return _strBlurb

            End Get

            Set(ByVal Value As String)

                _strBlurb = Value

            End Set

        End Property

        Public Sub setBlurb(ByVal Value As String)

            _strBlurb = Value

        End Sub
        Public Function getBlurb() As String

            Return _strBlurb

        End Function

        'Properties For Pagetext
        Public Property Pagetext() As String

            Get

                Return _strPagetext

            End Get

            Set(ByVal Value As String)

                _strPagetext = Value

            End Set

        End Property

        Public Sub setPagetext(ByVal Value As String)

            _strPagetext = Value

        End Sub
        Public Function getPagetext() As String

            Return _strPagetext

        End Function

        'Properties For UserId
        Public Property UserId() As Integer

            Get

                Return _intUserId

            End Get

            Set(ByVal Value As Integer)

                _intUserId = Value

            End Set

        End Property

        Public Sub setUserId(ByVal Value As Integer)

            _intUserId = Value

        End Sub
        Public Function getUserId() As Integer

            Return _intUserId

        End Function

        'Properties For LastModified
        Public Property LastModified() As Date

            Get

                Return _dtLastModified

            End Get

            Set(ByVal Value As Date)

                _dtLastModified = Value

            End Set

        End Property

        Public Sub setLastModified(ByVal Value As Date)

            _dtLastModified = Value

        End Sub
        Public Function getLastModified() As Date

            Return _dtLastModified

        End Function

        'Properties For ModBy
        Public Property ModBy() As Integer

            Get

                Return _intModBy

            End Get

            Set(ByVal Value As Integer)

                _intModBy = Value

            End Set

        End Property

        Public Sub setModBy(ByVal Value As Integer)

            _intModBy = Value

        End Sub
        Public Function getModBy() As Integer

            Return _intModBy

        End Function

        'Properties For ShowOnDroplist
        Public Property ShowOnDroplist() As Boolean

            Get

                Return _boolShowOnDroplist

            End Get

            Set(ByVal Value As Boolean)

                _boolShowOnDroplist = Value

            End Set

        End Property

        Public Sub setShowOnDroplist(ByVal Value As Boolean)

            _boolShowOnDroplist = Value

        End Sub
        Public Function getShowOnDroplist() As Boolean

            Return _boolShowOnDroplist

        End Function

        'Properties For ArticleCount
        Public Property ArticleCount() As Integer

            Get

                Return _intArticleCount

            End Get

            Set(ByVal Value As Integer)

                _intArticleCount = Value

            End Set

        End Property

        Public Sub setArticleCount(ByVal Value As Integer)

            _intArticleCount = Value

        End Sub
        Public Function getArticleCount() As Integer

            Return _intArticleCount

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intSectionId As Integer)

            Dim strCacheVar As String = "objCachedDbSections" & intSectionId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsByID " & intSectionId)

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

                    _intSectionId = objCached.getSectionId

                    _strSectionName = objCached.getSectionName

                    _intSectionParent = objCached.getSectionParent

                    _btPostInMenu = objCached.getPostInMenu

                    _strBlurb = objCached.getBlurb

                    _strPagetext = objCached.getPagetext

                    _intUserId = objCached.getUserId

                    _dtLastModified = objCached.getLastModified

                    _intModBy = objCached.getModBy

                    _boolShowOnDroplist = objCached.getShowOnDroplist

                    _intArticleCount = objCached.getArticleCount

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

                        _intSectionId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strSectionName = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intSectionParent = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _btPostInMenu = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strBlurb = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _strPagetext = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _intUserId = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _dtLastModified = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _intModBy = reader(8)

                    End If


                    If Not reader.IsDBNull(9) Then

                        _boolShowOnDroplist = reader(9)

                    End If


                    If Not reader.IsDBNull(10) Then

                        _intArticleCount = reader(10)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intSectionId As Integer)

            Dim strCacheVar As String = "objCachedDbSections" & intSectionId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteSections " & intSectionId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strSectionName As String, ByVal intSectionParent As Integer, ByVal btPostInMenu As Byte, ByVal strBlurb As String, ByVal strPagetext As String, ByVal intUserId As Integer, ByVal dtLastModified As Date, ByVal intModBy As Integer, ByVal boolShowOnDroplist As Boolean, ByVal intArticleCount As Integer)

            Dim intSectionId As Integer = 0

            Try

                intSectionId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertSections '" & strSectionName & "'," & intSectionParent & "," & btPostInMenu & ",'" & strBlurb & "','" & strPagetext & "'," & intUserId & ",'" & dtLastModified & "'," & intModBy & "," & boolShowOnDroplist & "," & intArticleCount))

                _intSectionId = intSectionId
                _strSectionName = strSectionName
                _intSectionParent = intSectionParent
                _btPostInMenu = btPostInMenu
                _strBlurb = strBlurb
                _strPagetext = strPagetext
                _intUserId = intUserId
                _dtLastModified = dtLastModified
                _intModBy = intModBy
                _boolShowOnDroplist = boolShowOnDroplist
                _intArticleCount = intArticleCount

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intSectionId As Integer = 0

            Try

                intSectionId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertSections '" & _strSectionName & "'," & _intSectionParent & "," & _btPostInMenu & ",'" & _strBlurb & "','" & _strPagetext & "'," & _intUserId & ",'" & _dtLastModified & "'," & _intModBy & "," & _boolShowOnDroplist & "," & _intArticleCount))

                _intSectionId = intSectionId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intSectionId As Integer, ByVal strSectionName As String, ByVal intSectionParent As Integer, ByVal btPostInMenu As Byte, ByVal strBlurb As String, ByVal strPagetext As String, ByVal intUserId As Integer, ByVal dtLastModified As Date, ByVal intModBy As Integer, ByVal boolShowOnDroplist As Boolean, ByVal intArticleCount As Integer)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateSections " & intSectionId & ",'" & strSectionName & "'," & intSectionParent & "," & btPostInMenu & ",'" & strBlurb & "','" & strPagetext & "'," & intUserId & ",'" & dtLastModified & "'," & intModBy & "," & boolShowOnDroplist & "," & intArticleCount)

                _intSectionId = intSectionId
                _strSectionName = strSectionName
                _intSectionParent = intSectionParent
                _btPostInMenu = btPostInMenu
                _strBlurb = strBlurb
                _strPagetext = strPagetext
                _intUserId = intUserId
                _dtLastModified = dtLastModified
                _intModBy = intModBy
                _boolShowOnDroplist = boolShowOnDroplist
                _intArticleCount = intArticleCount

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateSections " & _intSectionId & ",'" & _strSectionName & "'," & _intSectionParent & "," & _btPostInMenu & ",'" & _strBlurb & "','" & _strPagetext & "'," & _intUserId & ",'" & _dtLastModified & "'," & _intModBy & "," & _boolShowOnDroplist & "," & _intArticleCount)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.sections

            Dim strCacheVar As String = "objCachedDbSections" & _intSectionId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.sections

                objCached.setSectionId(_intSectionId)

                objCached.setSectionName(_strSectionName)

                objCached.setSectionParent(_intSectionParent)

                objCached.setPostInMenu(_btPostInMenu)

                objCached.setBlurb(_strBlurb)

                objCached.setPagetext(_strPagetext)

                objCached.setUserId(_intUserId)

                objCached.setLastModified(_dtLastModified)

                objCached.setModBy(_intModBy)

                objCached.setShowOnDroplist(_boolShowOnDroplist)

                objCached.setArticleCount(_intArticleCount)

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
