Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class Forums

        '*********************************************************************
        '  Class Name: Forums
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intForumid As Integer)


            If intForumid <> 0 Then

                getByID(intForumid)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intForumid As Integer = 0

        Private _strForum As String = ""

        Private _strDescription As String = ""

        Private _intPosts As Integer = 0

        Private _dtLastpost As Date = #1/1/1900#

#End Region

#Region "Properties"

        'Properties For Forumid
        Public Property Forumid() As Integer

            Get

                Return _intForumid

            End Get

            Set(ByVal Value As Integer)

                _intForumid = Value

            End Set

        End Property

        Public Sub setForumid(ByVal Value As Integer)

            _intForumid = Value

        End Sub
        Public Function getForumid() As Integer

            Return _intForumid

        End Function

        'Properties For Forum
        Public Property Forum() As String

            Get

                Return _strForum

            End Get

            Set(ByVal Value As String)

                _strForum = Value

            End Set

        End Property

        Public Sub setForum(ByVal Value As String)

            _strForum = Value

        End Sub
        Public Function getForum() As String

            Return _strForum

        End Function

        'Properties For Description
        Public Property Description() As String

            Get

                Return _strDescription

            End Get

            Set(ByVal Value As String)

                _strDescription = Value

            End Set

        End Property

        Public Sub setDescription(ByVal Value As String)

            _strDescription = Value

        End Sub
        Public Function getDescription() As String

            Return _strDescription

        End Function

        'Properties For Posts
        Public Property Posts() As Integer

            Get

                Return _intPosts

            End Get

            Set(ByVal Value As Integer)

                _intPosts = Value

            End Set

        End Property

        Public Sub setPosts(ByVal Value As Integer)

            _intPosts = Value

        End Sub
        Public Function getPosts() As Integer

            Return _intPosts

        End Function

        'Properties For Lastpost
        Public Property Lastpost() As Date

            Get

                Return _dtLastpost

            End Get

            Set(ByVal Value As Date)

                _dtLastpost = Value

            End Set

        End Property

        Public Sub setLastpost(ByVal Value As Date)

            _dtLastpost = Value

        End Sub
        Public Function getLastpost() As Date

            Return _dtLastpost

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intForumid As Integer)

            Dim strCacheVar As String = "objCachedDbForums" & intForumid

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetForumsByID " & intForumid)

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

                    _intForumid = objCached.getForumid

                    _strForum = objCached.getForum

                    _strDescription = objCached.getDescription

                    _intPosts = objCached.getPosts

                    _dtLastpost = objCached.getLastpost

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

                        _intForumid = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strForum = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _strDescription = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _intPosts = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _dtLastpost = reader(4)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intForumid As Integer)

            Dim strCacheVar As String = "objCachedDbForums" & intForumid

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteForums " & intForumid)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strForum As String, ByVal strDescription As String, ByVal intPosts As Integer, ByVal dtLastpost As Date)

            Dim intForumid As Integer = 0

            Try

                intForumid = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertForums '" & strForum & "','" & strDescription & "'," & intPosts & ",'" & dtLastpost & "'"))

                _intForumid = intForumid
                _strForum = strForum
                _strDescription = strDescription
                _intPosts = intPosts
                _dtLastpost = dtLastpost

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intForumid As Integer = 0

            Try

                intForumid = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertForums '" & _strForum & "','" & _strDescription & "'," & _intPosts & ",'" & _dtLastpost & "'"))

                _intForumid = intForumid

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intForumid As Integer, ByVal strForum As String, ByVal strDescription As String, ByVal intPosts As Integer, ByVal dtLastpost As Date)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateForums " & intForumid & ",'" & strForum & "','" & strDescription & "'," & intPosts & ",'" & dtLastpost & "'")

                _intForumid = intForumid
                _strForum = strForum
                _strDescription = strDescription
                _intPosts = intPosts
                _dtLastpost = dtLastpost

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateForums " & _intForumid & ",'" & _strForum & "','" & _strDescription & "'," & _intPosts & ",'" & _dtLastpost & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.Forums

            Dim strCacheVar As String = "objCachedDbForums" & _intForumid

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.Forums

                objCached.setForumid(_intForumid)

                objCached.setForum(_strForum)

                objCached.setDescription(_strDescription)

                objCached.setPosts(_intPosts)

                objCached.setLastpost(_dtLastpost)

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
