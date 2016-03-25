Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class cursos

        '*********************************************************************
        '  Class Name: cursos
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intCursoId As Integer)


            If intCursoId <> 0 Then

                getByID(intCursoId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intCursoId As Integer = 0

        Private _strTitle As String = ""

        Private _strCursoDesc As String = ""

        Private _strOnLine As String = ""

        Private _strZip As String = ""

        Private _intType As Integer = 0

#End Region

#Region "Properties"

        'Properties For CursoId
        Public Property CursoId() As Integer

            Get

                Return _intCursoId

            End Get

            Set(ByVal Value As Integer)

                _intCursoId = Value

            End Set

        End Property

        Public Sub setCursoId(ByVal Value As Integer)

            _intCursoId = Value

        End Sub
        Public Function getCursoId() As Integer

            Return _intCursoId

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

        'Properties For CursoDesc
        Public Property CursoDesc() As String

            Get

                Return _strCursoDesc

            End Get

            Set(ByVal Value As String)

                _strCursoDesc = Value

            End Set

        End Property

        Public Sub setCursoDesc(ByVal Value As String)

            _strCursoDesc = Value

        End Sub
        Public Function getCursoDesc() As String

            Return _strCursoDesc

        End Function

        'Properties For OnLine
        Public Property OnLine() As String

            Get

                Return _strOnLine

            End Get

            Set(ByVal Value As String)

                _strOnLine = Value

            End Set

        End Property

        Public Sub setOnLine(ByVal Value As String)

            _strOnLine = Value

        End Sub
        Public Function getOnLine() As String

            Return _strOnLine

        End Function

        'Properties For Zip
        Public Property Zip() As String

            Get

                Return _strZip

            End Get

            Set(ByVal Value As String)

                _strZip = Value

            End Set

        End Property

        Public Sub setZip(ByVal Value As String)

            _strZip = Value

        End Sub
        Public Function getZip() As String

            Return _strZip

        End Function

        'Properties For Type
        Public Property Type() As Integer

            Get

                Return _intType

            End Get

            Set(ByVal Value As Integer)

                _intType = Value

            End Set

        End Property

        Public Sub setCursoType(ByVal Value As Integer)

            _intType = Value

        End Sub
        Public Function getCursoType() As Integer

            Return _intType

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intCursoId As Integer)

            Dim strCacheVar As String = "objCachedDbCursos" & intCursoId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetCursosByID " & intCursoId)

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

                    _intCursoId = objCached.getCursoId

                    _strTitle = objCached.getTitle

                    _strCursoDesc = objCached.getCursoDesc

                    _strOnLine = objCached.getOnLine

                    _strZip = objCached.getZip

                    _intType = objCached.GetCursoType

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

                        _intCursoId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strTitle = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _strCursoDesc = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strOnLine = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strZip = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _intType = reader(5)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intCursoId As Integer)

            Dim strCacheVar As String = "objCachedDbCursos" & intCursoId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteCursos " & intCursoId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strTitle As String, ByVal strCursoDesc As String, ByVal strOnLine As String, ByVal strZip As String, ByVal intType As Integer)

            Dim intCursoId As Integer = 0

            Try

                intCursoId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertCursos '" & strTitle & "','" & strCursoDesc & "','" & strOnLine & "','" & strZip & "'," & intType))

                _intCursoId = intCursoId
                _strTitle = strTitle
                _strCursoDesc = strCursoDesc
                _strOnLine = strOnLine
                _strZip = strZip
                _intType = intType

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intCursoId As Integer = 0

            Try

                intCursoId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertCursos '" & _strTitle & "','" & _strCursoDesc & "','" & _strOnLine & "','" & _strZip & "'," & _intType))

                _intCursoId = intCursoId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intCursoId As Integer, ByVal strTitle As String, ByVal strCursoDesc As String, ByVal strOnLine As String, ByVal strZip As String, ByVal intType As Integer)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateCursos " & intCursoId & ",'" & strTitle & "','" & strCursoDesc & "','" & strOnLine & "','" & strZip & "'," & intType)

                _intCursoId = intCursoId
                _strTitle = strTitle
                _strCursoDesc = strCursoDesc
                _strOnLine = strOnLine
                _strZip = strZip
                _intType = intType

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateCursos " & _intCursoId & ",'" & _strTitle & "','" & _strCursoDesc & "','" & _strOnLine & "','" & _strZip & "'," & _intType)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.cursos

            Dim strCacheVar As String = "objCachedDbCursos" & _intCursoId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.cursos

                objCached.setCursoId(_intCursoId)

                objCached.setTitle(_strTitle)

                objCached.setCursoDesc(_strCursoDesc)

                objCached.setOnLine(_strOnLine)

                objCached.setZip(_strZip)

                objCached.setCursoType(_intType)

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
