Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class Messages

        '*********************************************************************
        '  Class Name: Messages
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intMessageid As Integer)


            If intMessageid <> 0 Then

                getByID(intMessageid)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intMessageid As Integer = 0

        Private _intForumid As Integer = 0

        Private _intInreplyto As Integer = 0

        Private _strPostedby As String = ""

        Private _strEmail As String = ""

        Private _strSubject As String = ""

        Private _strBody As String = ""

        Private _dtMdate As Date = #1/1/1900#

        Private _btApproved As Byte = 0

#End Region

#Region "Properties"

        'Properties For Messageid
        Public Property Messageid() As Integer

            Get

                Return _intMessageid

            End Get

            Set(ByVal Value As Integer)

                _intMessageid = Value

            End Set

        End Property

        Public Sub setMessageid(ByVal Value As Integer)

            _intMessageid = Value

        End Sub
        Public Function getMessageid() As Integer

            Return _intMessageid

        End Function

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

        'Properties For Inreplyto
        Public Property Inreplyto() As Integer

            Get

                Return _intInreplyto

            End Get

            Set(ByVal Value As Integer)

                _intInreplyto = Value

            End Set

        End Property

        Public Sub setInreplyto(ByVal Value As Integer)

            _intInreplyto = Value

        End Sub
        Public Function getInreplyto() As Integer

            Return _intInreplyto

        End Function

        'Properties For Postedby
        Public Property Postedby() As String

            Get

                Return _strPostedby

            End Get

            Set(ByVal Value As String)

                _strPostedby = Value

            End Set

        End Property

        Public Sub setPostedby(ByVal Value As String)

            _strPostedby = Value

        End Sub
        Public Function getPostedby() As String

            Return _strPostedby

        End Function

        'Properties For Email
        Public Property Email() As String

            Get

                Return _strEmail

            End Get

            Set(ByVal Value As String)

                _strEmail = Value

            End Set

        End Property

        Public Sub setEmail(ByVal Value As String)

            _strEmail = Value

        End Sub
        Public Function getEmail() As String

            Return _strEmail

        End Function

        'Properties For Subject
        Public Property Subject() As String

            Get

                Return _strSubject

            End Get

            Set(ByVal Value As String)

                _strSubject = Value

            End Set

        End Property

        Public Sub setSubject(ByVal Value As String)

            _strSubject = Value

        End Sub
        Public Function getSubject() As String

            Return _strSubject

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

        'Properties For Mdate
        Public Property Mdate() As Date

            Get

                Return _dtMdate

            End Get

            Set(ByVal Value As Date)

                _dtMdate = Value

            End Set

        End Property

        Public Sub setMdate(ByVal Value As Date)

            _dtMdate = Value

        End Sub
        Public Function getMdate() As Date

            Return _dtMdate

        End Function

        'Properties For Approved
        Public Property Approved() As Byte

            Get

                Return _btApproved

            End Get

            Set(ByVal Value As Byte)

                _btApproved = Value

            End Set

        End Property

        Public Sub setApproved(ByVal Value As Byte)

            _btApproved = Value

        End Sub
        Public Function getApproved() As Byte

            Return _btApproved

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intMessageid As Integer)

            Dim strCacheVar As String = "objCachedDbMessages" & intMessageid

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMessagesByID " & intMessageid)

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

                    _intMessageid = objCached.getMessageid

                    _intForumid = objCached.getForumid

                    _intInreplyto = objCached.getInreplyto

                    _strPostedby = objCached.getPostedby

                    _strEmail = objCached.getEmail

                    _strSubject = objCached.getSubject

                    _strBody = objCached.getBody

                    _dtMdate = objCached.getMdate

                    _btApproved = objCached.getApproved

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

                        _intMessageid = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _intForumid = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intInreplyto = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strPostedby = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strEmail = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _strSubject = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _strBody = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _dtMdate = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _btApproved = reader(8)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intMessageid As Integer)

            Dim strCacheVar As String = "objCachedDbMessages" & intMessageid

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteMessages " & intMessageid)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intForumid As Integer, ByVal intInreplyto As Integer, ByVal strPostedby As String, ByVal strEmail As String, ByVal strSubject As String, ByVal strBody As String, ByVal dtMdate As Date, ByVal btApproved As Byte)

            Dim intMessageid As Integer = 0

            Try

                intMessageid = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertMessages " & intForumid & "," & intInreplyto & ",'" & strPostedby & "','" & strEmail & "','" & strSubject & "','" & strBody & "','" & dtMdate & "'," & btApproved))

                _intMessageid = intMessageid
                _intForumid = intForumid
                _intInreplyto = intInreplyto
                _strPostedby = strPostedby
                _strEmail = strEmail
                _strSubject = strSubject
                _strBody = strBody
                _dtMdate = dtMdate
                _btApproved = btApproved

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intMessageid As Integer = 0

            Try

                intMessageid = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertMessages " & _intForumid & "," & _intInreplyto & ",'" & _strPostedby & "','" & _strEmail & "','" & _strSubject & "','" & _strBody & "','" & _dtMdate & "'," & _btApproved))

                _intMessageid = intMessageid

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intMessageid As Integer, ByVal intForumid As Integer, ByVal intInreplyto As Integer, ByVal strPostedby As String, ByVal strEmail As String, ByVal strSubject As String, ByVal strBody As String, ByVal dtMdate As Date, ByVal btApproved As Byte)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateMessages " & intMessageid & "," & intForumid & "," & intInreplyto & ",'" & strPostedby & "','" & strEmail & "','" & strSubject & "','" & strBody & "','" & dtMdate & "'," & btApproved)

                _intMessageid = intMessageid
                _intForumid = intForumid
                _intInreplyto = intInreplyto
                _strPostedby = strPostedby
                _strEmail = strEmail
                _strSubject = strSubject
                _strBody = strBody
                _dtMdate = dtMdate
                _btApproved = btApproved

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateMessages " & _intMessageid & "," & _intForumid & "," & _intInreplyto & ",'" & _strPostedby & "','" & _strEmail & "','" & _strSubject & "','" & _strBody & "','" & _dtMdate & "'," & _btApproved)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.Messages

            Dim strCacheVar As String = "objCachedDbMessages" & _intMessageid

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.Messages

                objCached.setMessageid(_intMessageid)

                objCached.setForumid(_intForumid)

                objCached.setInreplyto(_intInreplyto)

                objCached.setPostedby(_strPostedby)

                objCached.setEmail(_strEmail)

                objCached.setSubject(_strSubject)

                objCached.setBody(_strBody)

                objCached.setMdate(_dtMdate)

                objCached.setApproved(_btApproved)

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
