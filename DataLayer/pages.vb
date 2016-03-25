Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer

    Public Class pages

        '*********************************************************************
        '  Class Name: pages
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:50 PM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intPageId As Integer)


            If intPageId <> 0 Then

                getByID(intPageId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intPageId As Integer = 0

        Private _intSectionId As Integer = 0

        Private _dtPosted As Date = #1/1/1900#

        Private _strPageTitle As String = ""

        Private _strBlurb As String = ""

        Private _strBody As String = ""

        Private _btIsposted As Byte = 0

        Private _strPic As String = ""

        Private _intTypeofarticle As Integer = 0

        Private _btIsfeatured As Byte = 0

        Private _intUserId As Integer = 0

        Private _strKeywords As String = ""

        Private _intBook As Integer = 0

        Private _intChapter As Integer = 0

        Private _btRating As Byte = 0

#End Region

#Region "Properties"

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

        'Properties For Posted
        Public Property Posted() As Date

            Get

                Return _dtPosted

            End Get

            Set(ByVal Value As Date)

                _dtPosted = Value

            End Set

        End Property

        Public Sub setPosted(ByVal Value As Date)

            _dtPosted = Value

        End Sub
        Public Function getPosted() As Date

            Return _dtPosted

        End Function

        'Properties For PageTitle
        Public Property PageTitle() As String

            Get

                Return _strPageTitle

            End Get

            Set(ByVal Value As String)

                _strPageTitle = Value

            End Set

        End Property

        Public Sub setPageTitle(ByVal Value As String)

            _strPageTitle = Value

        End Sub
        Public Function getPageTitle() As String

            Return _strPageTitle

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

        'Properties For Isposted
        Public Property Isposted() As Byte

            Get

                Return _btIsposted

            End Get

            Set(ByVal Value As Byte)

                _btIsposted = Value

            End Set

        End Property

        Public Sub setIsposted(ByVal Value As Byte)

            _btIsposted = Value

        End Sub
        Public Function getIsposted() As Byte

            Return _btIsposted

        End Function

        'Properties For Pic
        Public Property Pic() As String

            Get

                Return _strPic

            End Get

            Set(ByVal Value As String)

                _strPic = Value

            End Set

        End Property

        Public Sub setPic(ByVal Value As String)

            _strPic = Value

        End Sub
        Public Function getPic() As String

            Return _strPic

        End Function

        'Properties For Typeofarticle
        Public Property Typeofarticle() As Integer

            Get

                Return _intTypeofarticle

            End Get

            Set(ByVal Value As Integer)

                _intTypeofarticle = Value

            End Set

        End Property

        Public Sub setTypeofarticle(ByVal Value As Integer)

            _intTypeofarticle = Value

        End Sub
        Public Function getTypeofarticle() As Integer

            Return _intTypeofarticle

        End Function

        'Properties For Isfeatured
        Public Property Isfeatured() As Byte

            Get

                Return _btIsfeatured

            End Get

            Set(ByVal Value As Byte)

                _btIsfeatured = Value

            End Set

        End Property

        Public Sub setIsfeatured(ByVal Value As Byte)

            _btIsfeatured = Value

        End Sub
        Public Function getIsfeatured() As Byte

            Return _btIsfeatured

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

        'Properties For Keywords
        Public Property Keywords() As String

            Get

                Return _strKeywords

            End Get

            Set(ByVal Value As String)

                _strKeywords = Value

            End Set

        End Property

        Public Sub setKeywords(ByVal Value As String)

            _strKeywords = Value

        End Sub
        Public Function getKeywords() As String

            Return _strKeywords

        End Function

        'Properties For Book
        Public Property Book() As Integer

            Get

                Return _intBook

            End Get

            Set(ByVal Value As Integer)

                _intBook = Value

            End Set

        End Property

        Public Sub setBook(ByVal Value As Integer)

            _intBook = Value

        End Sub
        Public Function getBook() As Integer

            Return _intBook

        End Function

        'Properties For Chapter
        Public Property Chapter() As Integer

            Get

                Return _intChapter

            End Get

            Set(ByVal Value As Integer)

                _intChapter = Value

            End Set

        End Property

        Public Sub setChapter(ByVal Value As Integer)

            _intChapter = Value

        End Sub
        Public Function getChapter() As Integer

            Return _intChapter

        End Function

        'Properties For Rating
        Public Property Rating() As Byte

            Get

                Return _btRating

            End Get

            Set(ByVal Value As Byte)

                _btRating = Value

            End Set

        End Property

        Public Sub setRating(ByVal Value As Byte)

            _btRating = Value

        End Sub
        Public Function getRating() As Byte

            Return _btRating

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intPageId As Integer)

            Dim strCacheVar As String = "objCachedDbPages" & intPageId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetPagesByID " & intPageId)

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

                    _intPageId = objCached.getPageId

                    _intSectionId = objCached.getSectionId

                    _dtPosted = objCached.getPosted

                    _strPageTitle = objCached.getPageTitle

                    _strBlurb = objCached.getBlurb

                    _strBody = objCached.getBody

                    _btIsposted = objCached.getIsposted

                    _strPic = objCached.getPic

                    _intTypeofarticle = objCached.getTypeofarticle

                    _btIsfeatured = objCached.getIsfeatured

                    _intUserId = objCached.getUserId

                    _strKeywords = objCached.getKeywords

                    _intBook = objCached.getBook

                    _intChapter = objCached.getChapter

                    _btRating = objCached.getRating

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

                        _intPageId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _intSectionId = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _dtPosted = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strPageTitle = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strBlurb = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _strBody = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _btIsposted = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _strPic = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _intTypeofarticle = reader(8)

                    End If


                    If Not reader.IsDBNull(9) Then

                        _btIsfeatured = reader(9)

                    End If


                    If Not reader.IsDBNull(10) Then

                        _intUserId = reader(10)

                    End If


                    If Not reader.IsDBNull(11) Then

                        _strKeywords = reader(11)

                    End If


                    If Not reader.IsDBNull(12) Then

                        _intBook = reader(12)

                    End If


                    If Not reader.IsDBNull(13) Then

                        _intChapter = reader(13)

                    End If


                    If Not reader.IsDBNull(14) Then

                        _btRating = reader(14)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intPageId As Integer)

            Dim strCacheVar As String = "objCachedDbPages" & intPageId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeletePages " & intPageId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal intSectionId As Integer, ByVal dtPosted As Date, ByVal strPageTitle As String, ByVal strBlurb As String, ByVal strBody As String, ByVal btIsposted As Byte, ByVal strPic As String, ByVal intTypeofarticle As Integer, ByVal btIsfeatured As Byte, ByVal intUserId As Integer, ByVal strKeywords As String, ByVal intBook As Integer, ByVal intChapter As Integer, ByVal btRating As Byte)

            Dim intPageId As Integer = 0

            Try

                intPageId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertPages " & intSectionId & ",'" & dtPosted & "','" & strPageTitle & "','" & strBlurb & "','" & strBody & "'," & btIsposted & ",'" & strPic & "'," & intTypeofarticle & "," & btIsfeatured & "," & intUserId & ",'" & strKeywords & "'," & intBook & "," & intChapter & "," & btRating))

                _intPageId = intPageId
                _intSectionId = intSectionId
                _dtPosted = dtPosted
                _strPageTitle = strPageTitle
                _strBlurb = strBlurb
                _strBody = strBody
                _btIsposted = btIsposted
                _strPic = strPic
                _intTypeofarticle = intTypeofarticle
                _btIsfeatured = btIsfeatured
                _intUserId = intUserId
                _strKeywords = strKeywords
                _intBook = intBook
                _intChapter = intChapter
                _btRating = btRating

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intPageId As Integer = 0

            Try

                intPageId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertPages " & _intSectionId & ",'" & _dtPosted & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating))

                _intPageId = intPageId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intPageId As Integer, ByVal intSectionId As Integer, ByVal dtPosted As Date, ByVal strPageTitle As String, ByVal strBlurb As String, ByVal strBody As String, ByVal btIsposted As Byte, ByVal strPic As String, ByVal intTypeofarticle As Integer, ByVal btIsfeatured As Byte, ByVal intUserId As Integer, ByVal strKeywords As String, ByVal intBook As Integer, ByVal intChapter As Integer, ByVal btRating As Byte)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdatePages " & intPageId & "," & intSectionId & ",'" & dtPosted & "','" & strPageTitle & "','" & strBlurb & "','" & strBody & "'," & btIsposted & ",'" & strPic & "'," & intTypeofarticle & "," & btIsfeatured & "," & intUserId & ",'" & strKeywords & "'," & intBook & "," & intChapter & "," & btRating)

                _intPageId = intPageId
                _intSectionId = intSectionId
                _dtPosted = dtPosted
                _strPageTitle = strPageTitle
                _strBlurb = strBlurb
                _strBody = strBody
                _btIsposted = btIsposted
                _strPic = strPic
                _intTypeofarticle = intTypeofarticle
                _btIsfeatured = btIsfeatured
                _intUserId = intUserId
                _strKeywords = strKeywords
                _intBook = intBook
                _intChapter = intChapter
                _btRating = btRating

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Function Update2() As String

            Try

                Dim strTxtFecha = _dtPosted.ToString()

                Dim anio As String '= cadena.substring(0, 4)
                Dim mes As String '= cadena.substring(3, 2)
                Dim dia As String '= cadena.substring(5, 2)
                dia = _dtPosted.Day.ToString
                mes = _dtPosted.Month.ToString
                anio = _dtPosted.Year.ToString

                'Dim strSplitFecha As String()

                'If strTxtFecha.IndexOf("/") > 0 Then
                '    strSplitFecha = strTxtFecha.Split("/")
                '    If strSplitFecha(0).Length = 2 Then
                '        'entonces esta colocando el dia primero
                '        dia = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        anio = strSplitFecha(2)
                '    End If
                '    If strSplitFecha(0).Length = 4 Then
                '        'entonces esta colocando el año primero
                '        anio = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        dia = strSplitFecha(2)
                '    End If
                'ElseIf strTxtFecha.IndexOf("-") > 0 Then
                '    strSplitFecha = strTxtFecha.Split("-")
                '    If strSplitFecha(0).Length = 2 Then
                '        'entonces esta colocando el dia primero
                '        dia = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        anio = strSplitFecha(2)
                '    End If
                '    If strSplitFecha(0).Length = 4 Then
                '        'entonces esta colocando el año primero
                '        anio = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        dia = strSplitFecha(2)
                '    End If
                'End If


                dia = agregaCeroaIzq(dia)
                mes = agregaCeroaIzq(mes)
                anio = anio.Substring(0, 4)
                'Dim dtNuevafecha As DateTime

                'dtNuevafecha = New DateTime(anio, mes, dia, 0, 0, 0)
                'Dim fecha As String = anio & "-" & mes & "-" & dia
                Dim fecha As String = anio & mes & dia

                'oObjeto.dtPosted.ToString("yyyyMMdd")

                'ParaLideres.GenericDataHandler.ExecSQL("proc_UpdatePages " & _intPageId & "," & _intSectionId & ",'" & _dtPosted & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating)

                Update2 = "PA_UpdatePages " & _intPageId & "," & _intSectionId & ",'" & fecha & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating

                'ParaLideres.GenericDataHandler.ExecSQL("PA_UpdatePages " & _intPageId & "," & _intSectionId & ",'" & fecha & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating)



                'UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Function

        Public Sub Update()

            Try

                Dim strTxtFecha = _dtPosted.ToString()

                Dim anio As String '= cadena.substring(0, 4)
                Dim mes As String '= cadena.substring(3, 2)
                Dim dia As String '= cadena.substring(5, 2)
                dia = _dtPosted.Day.ToString
                mes = _dtPosted.Month.ToString
                anio = _dtPosted.Year.ToString


                'Dim strSplitFecha As String()

                'If strTxtFecha.IndexOf("/") > 0 Then
                '    strSplitFecha = strTxtFecha.Split("/")
                '    If strSplitFecha(0).Length = 2 Then
                '        'entonces esta colocando el dia primero
                '        dia = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        anio = strSplitFecha(2)
                '    End If
                '    If strSplitFecha(0).Length = 4 Then
                '        'entonces esta colocando el año primero
                '        anio = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        dia = strSplitFecha(2)
                '    End If
                'ElseIf strTxtFecha.IndexOf("-") > 0 Then
                '    strSplitFecha = strTxtFecha.Split("-")
                '    If strSplitFecha(0).Length = 2 Then
                '        'entonces esta colocando el dia primero
                '        dia = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        anio = strSplitFecha(2)
                '    End If
                '    If strSplitFecha(0).Length = 4 Then
                '        'entonces esta colocando el año primero
                '        anio = strSplitFecha(0)
                '        mes = strSplitFecha(1)
                '        dia = strSplitFecha(2)
                '    End If
                'End If


                dia = agregaCeroaIzq(dia)
                mes = agregaCeroaIzq(mes)
                anio = anio.Substring(0, 4)
                'Dim dtNuevafecha As DateTime

                'dtNuevafecha = New DateTime(anio, mes, dia, 0, 0, 0)
                'Dim fecha As String = anio & "-" & mes & "-" & dia
                Dim fecha As String = anio & mes & dia

                'oObjeto.dtPosted.ToString("yyyyMMdd")

                'ParaLideres.GenericDataHandler.ExecSQL("proc_UpdatePages " & _intPageId & "," & _intSectionId & ",'" & _dtPosted & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating)


                ParaLideres.GenericDataHandler.ExecSQL("PA_UpdatePages " & _intPageId & "," & _intSectionId & ",'" & fecha & "','" & _strPageTitle & "','" & _strBlurb & "','" & _strBody & "'," & _btIsposted & ",'" & _strPic & "'," & _intTypeofarticle & "," & _btIsfeatured & "," & _intUserId & ",'" & _strKeywords & "'," & _intBook & "," & _intChapter & "," & _btRating)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        Function agregaCeroaIzq(ByVal texto As String) As String
            Dim nuevotexto As String = ""
            If texto.Length = 1 Then
                nuevotexto = "0" & texto
            Else
                nuevotexto = texto
            End If
            Return nuevotexto
        End Function

        Public Sub UpdateCache()

            Dim objCached As DataLayer.pages

            Dim strCacheVar As String = "objCachedDbPages" & _intPageId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.pages

                objCached.setPageId(_intPageId)

                objCached.setSectionId(_intSectionId)

                objCached.setPosted(_dtPosted)

                objCached.setPageTitle(_strPageTitle)

                objCached.setBlurb(_strBlurb)

                objCached.setBody(_strBody)

                objCached.setIsposted(_btIsposted)

                objCached.setPic(_strPic)

                objCached.setTypeofarticle(_intTypeofarticle)

                objCached.setIsfeatured(_btIsfeatured)

                objCached.setUserId(_intUserId)

                objCached.setKeywords(_strKeywords)

                objCached.setBook(_intBook)

                objCached.setChapter(_intChapter)

                objCached.setRating(_btRating)

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
