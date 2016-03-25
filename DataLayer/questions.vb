Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class questions

        '*********************************************************************
        '  Class Name: questions
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:50 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intQuestionId As Integer)


            If intQuestionId <> 0 Then

                getByID(intQuestionId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intQuestionId As Integer = 0

        Private _strQuestionDesc As String = ""

        Private _dtDateStart As Date = #1/1/1900#

        Private _dtDateEnd As Date = #1/1/1900#

        Private _strQuestion1 As String = ""

        Private _strQuestion2 As String = ""

        Private _strQuestion3 As String = ""

        Private _strQuestion4 As String = ""

        Private _strQuestion5 As String = ""

        Private _strQuestion6 As String = ""

        Private _strQuestion7 As String = ""

        Private _strQuestion8 As String = ""

        Private _strQuestion9 As String = ""

        Private _strQuestion10 As String = ""

#End Region

#Region "Properties"

        'Properties For QuestionId
        Public Property QuestionId() As Integer

            Get

                Return _intQuestionId

            End Get

            Set(ByVal Value As Integer)

                _intQuestionId = Value

            End Set

        End Property

        Public Sub setQuestionId(ByVal Value As Integer)

            _intQuestionId = Value

        End Sub
        Public Function getQuestionId() As Integer

            Return _intQuestionId

        End Function

        'Properties For QuestionDesc
        Public Property QuestionDesc() As String

            Get

                Return _strQuestionDesc

            End Get

            Set(ByVal Value As String)

                _strQuestionDesc = Value

            End Set

        End Property

        Public Sub setQuestionDesc(ByVal Value As String)

            _strQuestionDesc = Value

        End Sub
        Public Function getQuestionDesc() As String

            Return _strQuestionDesc

        End Function

        'Properties For DateStart
        Public Property DateStart() As Date

            Get

                Return _dtDateStart

            End Get

            Set(ByVal Value As Date)

                _dtDateStart = Value

            End Set

        End Property

        Public Sub setDateStart(ByVal Value As Date)

            _dtDateStart = Value

        End Sub
        Public Function getDateStart() As Date

            Return _dtDateStart

        End Function

        'Properties For DateEnd
        Public Property DateEnd() As Date

            Get

                Return _dtDateEnd

            End Get

            Set(ByVal Value As Date)

                _dtDateEnd = Value

            End Set

        End Property

        Public Sub setDateEnd(ByVal Value As Date)

            _dtDateEnd = Value

        End Sub
        Public Function getDateEnd() As Date

            Return _dtDateEnd

        End Function

        'Properties For Question1
        Public Property Question1() As String

            Get

                Return _strQuestion1

            End Get

            Set(ByVal Value As String)

                _strQuestion1 = Value

            End Set

        End Property

        Public Sub setQuestion1(ByVal Value As String)

            _strQuestion1 = Value

        End Sub
        Public Function getQuestion1() As String

            Return _strQuestion1

        End Function

        'Properties For Question2
        Public Property Question2() As String

            Get

                Return _strQuestion2

            End Get

            Set(ByVal Value As String)

                _strQuestion2 = Value

            End Set

        End Property

        Public Sub setQuestion2(ByVal Value As String)

            _strQuestion2 = Value

        End Sub
        Public Function getQuestion2() As String

            Return _strQuestion2

        End Function

        'Properties For Question3
        Public Property Question3() As String

            Get

                Return _strQuestion3

            End Get

            Set(ByVal Value As String)

                _strQuestion3 = Value

            End Set

        End Property

        Public Sub setQuestion3(ByVal Value As String)

            _strQuestion3 = Value

        End Sub
        Public Function getQuestion3() As String

            Return _strQuestion3

        End Function

        'Properties For Question4
        Public Property Question4() As String

            Get

                Return _strQuestion4

            End Get

            Set(ByVal Value As String)

                _strQuestion4 = Value

            End Set

        End Property

        Public Sub setQuestion4(ByVal Value As String)

            _strQuestion4 = Value

        End Sub
        Public Function getQuestion4() As String

            Return _strQuestion4

        End Function

        'Properties For Question5
        Public Property Question5() As String

            Get

                Return _strQuestion5

            End Get

            Set(ByVal Value As String)

                _strQuestion5 = Value

            End Set

        End Property

        Public Sub setQuestion5(ByVal Value As String)

            _strQuestion5 = Value

        End Sub
        Public Function getQuestion5() As String

            Return _strQuestion5

        End Function

        'Properties For Question6
        Public Property Question6() As String

            Get

                Return _strQuestion6

            End Get

            Set(ByVal Value As String)

                _strQuestion6 = Value

            End Set

        End Property

        Public Sub setQuestion6(ByVal Value As String)

            _strQuestion6 = Value

        End Sub
        Public Function getQuestion6() As String

            Return _strQuestion6

        End Function

        'Properties For Question7
        Public Property Question7() As String

            Get

                Return _strQuestion7

            End Get

            Set(ByVal Value As String)

                _strQuestion7 = Value

            End Set

        End Property

        Public Sub setQuestion7(ByVal Value As String)

            _strQuestion7 = Value

        End Sub
        Public Function getQuestion7() As String

            Return _strQuestion7

        End Function

        'Properties For Question8
        Public Property Question8() As String

            Get

                Return _strQuestion8

            End Get

            Set(ByVal Value As String)

                _strQuestion8 = Value

            End Set

        End Property

        Public Sub setQuestion8(ByVal Value As String)

            _strQuestion8 = Value

        End Sub
        Public Function getQuestion8() As String

            Return _strQuestion8

        End Function

        'Properties For Question9
        Public Property Question9() As String

            Get

                Return _strQuestion9

            End Get

            Set(ByVal Value As String)

                _strQuestion9 = Value

            End Set

        End Property

        Public Sub setQuestion9(ByVal Value As String)

            _strQuestion9 = Value

        End Sub
        Public Function getQuestion9() As String

            Return _strQuestion9

        End Function

        'Properties For Question10
        Public Property Question10() As String

            Get

                Return _strQuestion10

            End Get

            Set(ByVal Value As String)

                _strQuestion10 = Value

            End Set

        End Property

        Public Sub setQuestion10(ByVal Value As String)

            _strQuestion10 = Value

        End Sub
        Public Function getQuestion10() As String

            Return _strQuestion10

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intQuestionId As Integer)

            Dim strCacheVar As String = "objCachedDbQuestions" & intQuestionId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetQuestionsByID " & intQuestionId)

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

                    _intQuestionId = objCached.getQuestionId

                    _strQuestionDesc = objCached.getQuestionDesc

                    _dtDateStart = objCached.getDateStart

                    _dtDateEnd = objCached.getDateEnd

                    _strQuestion1 = objCached.getQuestion1

                    _strQuestion2 = objCached.getQuestion2

                    _strQuestion3 = objCached.getQuestion3

                    _strQuestion4 = objCached.getQuestion4

                    _strQuestion5 = objCached.getQuestion5

                    _strQuestion6 = objCached.getQuestion6

                    _strQuestion7 = objCached.getQuestion7

                    _strQuestion8 = objCached.getQuestion8

                    _strQuestion9 = objCached.getQuestion9

                    _strQuestion10 = objCached.getQuestion10

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

                        _intQuestionId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strQuestionDesc = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _dtDateStart = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _dtDateEnd = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strQuestion1 = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _strQuestion2 = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _strQuestion3 = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _strQuestion4 = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _strQuestion5 = reader(8)

                    End If


                    If Not reader.IsDBNull(9) Then

                        _strQuestion6 = reader(9)

                    End If


                    If Not reader.IsDBNull(10) Then

                        _strQuestion7 = reader(10)

                    End If


                    If Not reader.IsDBNull(11) Then

                        _strQuestion8 = reader(11)

                    End If


                    If Not reader.IsDBNull(12) Then

                        _strQuestion9 = reader(12)

                    End If


                    If Not reader.IsDBNull(13) Then

                        _strQuestion10 = reader(13)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intQuestionId As Integer)

            Dim strCacheVar As String = "objCachedDbQuestions" & intQuestionId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteQuestions " & intQuestionId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strQuestionDesc As String, ByVal dtDateStart As Date, ByVal dtDateEnd As Date, ByVal strQuestion1 As String, ByVal strQuestion2 As String, ByVal strQuestion3 As String, ByVal strQuestion4 As String, ByVal strQuestion5 As String, ByVal strQuestion6 As String, ByVal strQuestion7 As String, ByVal strQuestion8 As String, ByVal strQuestion9 As String, ByVal strQuestion10 As String)

            Dim intQuestionId As Integer = 0

            Try

                intQuestionId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertQuestions '" & strQuestionDesc & "','" & dtDateStart & "','" & dtDateEnd & "','" & strQuestion1 & "','" & strQuestion2 & "','" & strQuestion3 & "','" & strQuestion4 & "','" & strQuestion5 & "','" & strQuestion6 & "','" & strQuestion7 & "','" & strQuestion8 & "','" & strQuestion9 & "','" & strQuestion10 & "'"))

                _intQuestionId = intQuestionId
                _strQuestionDesc = strQuestionDesc
                _dtDateStart = dtDateStart
                _dtDateEnd = dtDateEnd
                _strQuestion1 = strQuestion1
                _strQuestion2 = strQuestion2
                _strQuestion3 = strQuestion3
                _strQuestion4 = strQuestion4
                _strQuestion5 = strQuestion5
                _strQuestion6 = strQuestion6
                _strQuestion7 = strQuestion7
                _strQuestion8 = strQuestion8
                _strQuestion9 = strQuestion9
                _strQuestion10 = strQuestion10

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intQuestionId As Integer = 0

            Try

                intQuestionId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertQuestions '" & _strQuestionDesc & "','" & _dtDateStart & "','" & _dtDateEnd & "','" & _strQuestion1 & "','" & _strQuestion2 & "','" & _strQuestion3 & "','" & _strQuestion4 & "','" & _strQuestion5 & "','" & _strQuestion6 & "','" & _strQuestion7 & "','" & _strQuestion8 & "','" & _strQuestion9 & "','" & _strQuestion10 & "'"))

                _intQuestionId = intQuestionId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intQuestionId As Integer, ByVal strQuestionDesc As String, ByVal dtDateStart As Date, ByVal dtDateEnd As Date, ByVal strQuestion1 As String, ByVal strQuestion2 As String, ByVal strQuestion3 As String, ByVal strQuestion4 As String, ByVal strQuestion5 As String, ByVal strQuestion6 As String, ByVal strQuestion7 As String, ByVal strQuestion8 As String, ByVal strQuestion9 As String, ByVal strQuestion10 As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateQuestions " & intQuestionId & ",'" & strQuestionDesc & "','" & dtDateStart & "','" & dtDateEnd & "','" & strQuestion1 & "','" & strQuestion2 & "','" & strQuestion3 & "','" & strQuestion4 & "','" & strQuestion5 & "','" & strQuestion6 & "','" & strQuestion7 & "','" & strQuestion8 & "','" & strQuestion9 & "','" & strQuestion10 & "'")

                _intQuestionId = intQuestionId
                _strQuestionDesc = strQuestionDesc
                _dtDateStart = dtDateStart
                _dtDateEnd = dtDateEnd
                _strQuestion1 = strQuestion1
                _strQuestion2 = strQuestion2
                _strQuestion3 = strQuestion3
                _strQuestion4 = strQuestion4
                _strQuestion5 = strQuestion5
                _strQuestion6 = strQuestion6
                _strQuestion7 = strQuestion7
                _strQuestion8 = strQuestion8
                _strQuestion9 = strQuestion9
                _strQuestion10 = strQuestion10

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateQuestions " & _intQuestionId & ",'" & _strQuestionDesc & "','" & _dtDateStart & "','" & _dtDateEnd & "','" & _strQuestion1 & "','" & _strQuestion2 & "','" & _strQuestion3 & "','" & _strQuestion4 & "','" & _strQuestion5 & "','" & _strQuestion6 & "','" & _strQuestion7 & "','" & _strQuestion8 & "','" & _strQuestion9 & "','" & _strQuestion10 & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.questions

            Dim strCacheVar As String = "objCachedDbQuestions" & _intQuestionId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.questions

                objCached.setQuestionId(_intQuestionId)

                objCached.setQuestionDesc(_strQuestionDesc)

                objCached.setDateStart(_dtDateStart)

                objCached.setDateEnd(_dtDateEnd)

                objCached.setQuestion1(_strQuestion1)

                objCached.setQuestion2(_strQuestion2)

                objCached.setQuestion3(_strQuestion3)

                objCached.setQuestion4(_strQuestion4)

                objCached.setQuestion5(_strQuestion5)

                objCached.setQuestion6(_strQuestion6)

                objCached.setQuestion7(_strQuestion7)

                objCached.setQuestion8(_strQuestion8)

                objCached.setQuestion9(_strQuestion9)

                objCached.setQuestion10(_strQuestion10)

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
