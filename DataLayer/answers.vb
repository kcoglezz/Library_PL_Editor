Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class answers

        '*********************************************************************
        '  Class Name: answers
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intAnswerId As Integer)


            If intAnswerId <> 0 Then

                getByID(intAnswerId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intAnswerId As Integer = 0

        Private _intQuestionId As Integer = 0

        Private _intVal1 As Integer = 0

        Private _intVal2 As Integer = 0

        Private _intVal3 As Integer = 0

        Private _intVal4 As Integer = 0

        Private _intVal5 As Integer = 0

        Private _intVal6 As Integer = 0

        Private _intVal7 As Integer = 0

        Private _intVal8 As Integer = 0

        Private _intVal9 As Integer = 0

        Private _intVal10 As Integer = 0

        Private _strIpaddress As String = ""

#End Region

#Region "Properties"

        'Properties For AnswerId
        Public Property AnswerId() As Integer

            Get

                Return _intAnswerId

            End Get

            Set(ByVal Value As Integer)

                _intAnswerId = Value

            End Set

        End Property

        Public Sub setAnswerId(ByVal Value As Integer)

            _intAnswerId = Value

        End Sub
        Public Function getAnswerId() As Integer

            Return _intAnswerId

        End Function

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

        'Properties For Val1
        Public Property Val1() As Integer

            Get

                Return _intVal1

            End Get

            Set(ByVal Value As Integer)

                _intVal1 = Value

            End Set

        End Property

        Public Sub setVal1(ByVal Value As Integer)

            _intVal1 = Value

        End Sub
        Public Function getVal1() As Integer

            Return _intVal1

        End Function

        'Properties For Val2
        Public Property Val2() As Integer

            Get

                Return _intVal2

            End Get

            Set(ByVal Value As Integer)

                _intVal2 = Value

            End Set

        End Property

        Public Sub setVal2(ByVal Value As Integer)

            _intVal2 = Value

        End Sub
        Public Function getVal2() As Integer

            Return _intVal2

        End Function

        'Properties For Val3
        Public Property Val3() As Integer

            Get

                Return _intVal3

            End Get

            Set(ByVal Value As Integer)

                _intVal3 = Value

            End Set

        End Property

        Public Sub setVal3(ByVal Value As Integer)

            _intVal3 = Value

        End Sub
        Public Function getVal3() As Integer

            Return _intVal3

        End Function

        'Properties For Val4
        Public Property Val4() As Integer

            Get

                Return _intVal4

            End Get

            Set(ByVal Value As Integer)

                _intVal4 = Value

            End Set

        End Property

        Public Sub setVal4(ByVal Value As Integer)

            _intVal4 = Value

        End Sub
        Public Function getVal4() As Integer

            Return _intVal4

        End Function

        'Properties For Val5
        Public Property Val5() As Integer

            Get

                Return _intVal5

            End Get

            Set(ByVal Value As Integer)

                _intVal5 = Value

            End Set

        End Property

        Public Sub setVal5(ByVal Value As Integer)

            _intVal5 = Value

        End Sub
        Public Function getVal5() As Integer

            Return _intVal5

        End Function

        'Properties For Val6
        Public Property Val6() As Integer

            Get

                Return _intVal6

            End Get

            Set(ByVal Value As Integer)

                _intVal6 = Value

            End Set

        End Property

        Public Sub setVal6(ByVal Value As Integer)

            _intVal6 = Value

        End Sub
        Public Function getVal6() As Integer

            Return _intVal6

        End Function

        'Properties For Val7
        Public Property Val7() As Integer

            Get

                Return _intVal7

            End Get

            Set(ByVal Value As Integer)

                _intVal7 = Value

            End Set

        End Property

        Public Sub setVal7(ByVal Value As Integer)

            _intVal7 = Value

        End Sub
        Public Function getVal7() As Integer

            Return _intVal7

        End Function

        'Properties For Val8
        Public Property Val8() As Integer

            Get

                Return _intVal8

            End Get

            Set(ByVal Value As Integer)

                _intVal8 = Value

            End Set

        End Property

        Public Sub setVal8(ByVal Value As Integer)

            _intVal8 = Value

        End Sub
        Public Function getVal8() As Integer

            Return _intVal8

        End Function

        'Properties For Val9
        Public Property Val9() As Integer

            Get

                Return _intVal9

            End Get

            Set(ByVal Value As Integer)

                _intVal9 = Value

            End Set

        End Property

        Public Sub setVal9(ByVal Value As Integer)

            _intVal9 = Value

        End Sub
        Public Function getVal9() As Integer

            Return _intVal9

        End Function

        'Properties For Val10
        Public Property Val10() As Integer

            Get

                Return _intVal10

            End Get

            Set(ByVal Value As Integer)

                _intVal10 = Value

            End Set

        End Property

        Public Sub setVal10(ByVal Value As Integer)

            _intVal10 = Value

        End Sub
        Public Function getVal10() As Integer

            Return _intVal10

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
        Public Sub getByID(ByVal intAnswerId As Integer)

            Dim strCacheVar As String = "objCachedDbAnswers" & intAnswerId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetAnswersByID " & intAnswerId)

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

                    _intAnswerId = objCached.getAnswerId

                    _intQuestionId = objCached.getQuestionId

                    _intVal1 = objCached.getVal1

                    _intVal2 = objCached.getVal2

                    _intVal3 = objCached.getVal3

                    _intVal4 = objCached.getVal4

                    _intVal5 = objCached.getVal5

                    _intVal6 = objCached.getVal6

                    _intVal7 = objCached.getVal7

                    _intVal8 = objCached.getVal8

                    _intVal9 = objCached.getVal9

                    _intVal10 = objCached.getVal10

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

                        _intAnswerId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _intQuestionId = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _intVal1 = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _intVal2 = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _intVal3 = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _intVal4 = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _intVal5 = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _intVal6 = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _intVal7 = reader(8)

                    End If


                    If Not reader.IsDBNull(9) Then

                        _intVal8 = reader(9)

                    End If


                    If Not reader.IsDBNull(10) Then

                        _intVal9 = reader(10)

                    End If


                    If Not reader.IsDBNull(11) Then

                        _intVal10 = reader(11)

                    End If


                    If Not reader.IsDBNull(12) Then

                        _strIpaddress = reader(12)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intAnswerId As Integer)

            Dim strCacheVar As String = "objCachedDbAnswers" & intAnswerId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteAnswers " & intAnswerId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intQuestionId As Integer, ByVal intVal1 As Integer, ByVal intVal2 As Integer, ByVal intVal3 As Integer, ByVal intVal4 As Integer, ByVal intVal5 As Integer, ByVal intVal6 As Integer, ByVal intVal7 As Integer, ByVal intVal8 As Integer, ByVal intVal9 As Integer, ByVal intVal10 As Integer, ByVal strIpaddress As String)

            Dim intAnswerId As Integer = 0

            Try

                intAnswerId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertAnswers " & intQuestionId & "," & intVal1 & "," & intVal2 & "," & intVal3 & "," & intVal4 & "," & intVal5 & "," & intVal6 & "," & intVal7 & "," & intVal8 & "," & intVal9 & "," & intVal10 & ",'" & strIpaddress & "'"))

                _intAnswerId = intAnswerId
                _intQuestionId = intQuestionId
                _intVal1 = intVal1
                _intVal2 = intVal2
                _intVal3 = intVal3
                _intVal4 = intVal4
                _intVal5 = intVal5
                _intVal6 = intVal6
                _intVal7 = intVal7
                _intVal8 = intVal8
                _intVal9 = intVal9
                _intVal10 = intVal10
                _strIpaddress = strIpaddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intAnswerId As Integer = 0

            Try

                intAnswerId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertAnswers " & _intQuestionId & "," & _intVal1 & "," & _intVal2 & "," & _intVal3 & "," & _intVal4 & "," & _intVal5 & "," & _intVal6 & "," & _intVal7 & "," & _intVal8 & "," & _intVal9 & "," & _intVal10 & ",'" & _strIpaddress & "'"))

                _intAnswerId = intAnswerId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intAnswerId As Integer, ByVal intQuestionId As Integer, ByVal intVal1 As Integer, ByVal intVal2 As Integer, ByVal intVal3 As Integer, ByVal intVal4 As Integer, ByVal intVal5 As Integer, ByVal intVal6 As Integer, ByVal intVal7 As Integer, ByVal intVal8 As Integer, ByVal intVal9 As Integer, ByVal intVal10 As Integer, ByVal strIpaddress As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateAnswers " & intAnswerId & "," & intQuestionId & "," & intVal1 & "," & intVal2 & "," & intVal3 & "," & intVal4 & "," & intVal5 & "," & intVal6 & "," & intVal7 & "," & intVal8 & "," & intVal9 & "," & intVal10 & ",'" & strIpaddress & "'")

                _intAnswerId = intAnswerId
                _intQuestionId = intQuestionId
                _intVal1 = intVal1
                _intVal2 = intVal2
                _intVal3 = intVal3
                _intVal4 = intVal4
                _intVal5 = intVal5
                _intVal6 = intVal6
                _intVal7 = intVal7
                _intVal8 = intVal8
                _intVal9 = intVal9
                _intVal10 = intVal10
                _strIpaddress = strIpaddress

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateAnswers " & _intAnswerId & "," & _intQuestionId & "," & _intVal1 & "," & _intVal2 & "," & _intVal3 & "," & _intVal4 & "," & _intVal5 & "," & _intVal6 & "," & _intVal7 & "," & _intVal8 & "," & _intVal9 & "," & _intVal10 & ",'" & _strIpaddress & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.answers

            Dim strCacheVar As String = "objCachedDbAnswers" & _intAnswerId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.answers

                objCached.setAnswerId(_intAnswerId)

                objCached.setQuestionId(_intQuestionId)

                objCached.setVal1(_intVal1)

                objCached.setVal2(_intVal2)

                objCached.setVal3(_intVal3)

                objCached.setVal4(_intVal4)

                objCached.setVal5(_intVal5)

                objCached.setVal6(_intVal6)

                objCached.setVal7(_intVal7)

                objCached.setVal8(_intVal8)

                objCached.setVal9(_intVal9)

                objCached.setVal10(_intVal10)

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
