Imports System.Text
Imports System.Web
Imports Sistema.PL.Entidad

Namespace ParaLideres

    Public Class Functions

        Private Shared _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath_editor")
        'Private Shared _mail_server As String = System.Web.Configuration.WebConfigurationManager.AppSettings("MailServer")
        Private Shared _support_account As String = System.Web.Configuration.WebConfigurationManager.AppSettings("SupportAccount")
        Private Shared _contact_account As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ContactAccount")
        Private Shared _develop As String = System.Web.Configuration.WebConfigurationManager.AppSettings("DeveloperAccount")

        'Private Shared strServer As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Server")
        'Private Shared strDatabase As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Database")
        'Private Shared strUsername As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Username")
        'Private Shared strPassword As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Password")
        'Private Shared strConnection As String = "User ID=" & strUsername & ";Password=" & strPassword & ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer & ";Network Library =dbmssocn"

        Private Shared oConex As New InfoConexion()
        Private Shared strConnection As String = oConex.StringConnection
        Private Shared _strGeoIPDataFilePath As String = HttpContext.Current.Server.MapPath("GeoData\GeoIP.dat")
        Private Shared _strBorderColor As String = "#D09000" 'orange
        Private Shared _strBackGroundColor As String = "#A7D168" 'green

        Private Shared _intWidth As Integer = 450



        Public Enum ColType As Integer

            Text = 1
            Number = 2
            Money = 3
            Percentage = 4
            DateType = 5
            TextAlignCenter = 6
            TextAlignRight = 7
            ImageCheck = 8

        End Enum


        Public Shared ReadOnly Property Width() As Integer
            Get
                Return _intWidth

            End Get
        End Property

        Public Shared ReadOnly Property SupportAccount() As String
            Get
                Return _support_account
            End Get
        End Property

        Public Shared ReadOnly Property DevelopAccount() As String
            Get
                Return _develop
            End Get
        End Property

        Public Shared ReadOnly Property Contact() As String
            Get

                Return _contact_account

            End Get
        End Property


        Public Shared ReadOnly Property ProjectPath() As String
            Get
                Return _project_path
            End Get
        End Property

        Public Shared ReadOnly Property GeoDataFilePath() As String
            Get
                Return _strGeoIPDataFilePath

            End Get
        End Property

        Public Shared Function CheckDateFormat(ByVal strDate As String) As Date

            Dim arrVals As String() = Split(strDate, "/")
            Dim dtValue As Date = #1/1/1900#

            Try

            Catch ex As Exception

            End Try

        End Function


        Public Shared Function GetCountryCode() As String

            Dim strCountryCode As String = ""

            Try

                If IsNothing(HttpContext.Current.Session("country_code")) Or HttpContext.Current.Session("country_code") = "" Then

                    Dim objCountryLookup As CountryLookup

                    Dim strUserIPAddress As String = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")

                    Try

                        objCountryLookup = New CountryLookup(_strGeoIPDataFilePath)

                        strCountryCode = objCountryLookup.LookupCountryCode(strUserIPAddress)

                        Try

                            HttpContext.Current.Session("country_id") = CInt(ParaLideres.GenericDataHandler.ExecScalar("SELECT id FROM countries WHERE country_code = '" & strCountryCode & "'"))

                        Catch ex As Exception

                            HttpContext.Current.Session("country_id") = 0

                        End Try


                        HttpContext.Current.Session("country_code") = strCountryCode

                    Catch ex As Exception

                        HttpContext.Current.Trace.Warn(ex.ToString())

                    Finally

                        objCountryLookup = Nothing

                    End Try

                Else 'If Session("country_code") = "" Then

                    strCountryCode = HttpContext.Current.Session("country_code")

                End If 'If Session("country_code") = "" Then

                Return strCountryCode

            Catch ex As Exception

                Throw ex

            Finally

            End Try

        End Function

        Public Shared Function SpanishMonthName(ByVal intMonth As Integer) As String

            Dim strMonthName As String = ""

            Try

                Select Case intMonth

                    Case 1

                        strMonthName = "Enero"

                    Case 2

                        strMonthName = "Febrero"

                    Case 3

                        strMonthName = "Marzo"

                    Case 4

                        strMonthName = "Abril"

                    Case 5

                        strMonthName = "Mayo"

                    Case 6

                        strMonthName = "Junio"

                    Case 7

                        strMonthName = "Julio"

                    Case 8

                        strMonthName = "Agosto"

                    Case 9

                        strMonthName = "Septiembre"

                    Case 10

                        strMonthName = "Octubre"

                    Case 11

                        strMonthName = "Noviembre"

                    Case 12

                        strMonthName = "Diciembre"

                End Select

                Return strMonthName

            Catch ex As Exception

                Throw ex

            End Try

        End Function


        Public Shared Function FormatHispanicDateTime(ByVal dtValue As Date) As String

            Dim sb As New StringBuilder("")

            Try

                sb.Append(Day(dtValue) & "/" & EspMonth(Month(dtValue)) & "/" & Year(dtValue))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Shared Function FormatHispanicDateTime(ByVal strDate As String) As String

            Dim sb As New StringBuilder("")

            Dim dtValue As Date

            Try

                dtValue = CDate(strDate)

                sb.Append(Day(dtValue) & "/" & EspMonth(Month(dtValue)) & "/" & Year(dtValue))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function FormatHispanicDateTimeShort(ByVal strDate As String) As String

            Dim sb As New StringBuilder("")

            Dim dtValue As Date

            Try

                dtValue = CDate(strDate)

                sb.Append(Day(dtValue) & "/" & Month(dtValue) & "/" & Year(dtValue))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function FormatHispanicDateToEnglish(ByVal strDate As String) As String

            Dim sb As New StringBuilder("")

            Dim arrVals As String() = Split(strDate, "/")

            Try

                sb.Append(arrVals(1) & "/" & arrVals(0) & "/" & arrVals(2))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function FormatRequestedDate(ByVal strFrmFldName As String) As Date

            Dim dtResult As Date = #1/1/1900#

            Dim intMonth As Integer
            Dim intDay As Integer
            Dim intYear As Integer

            Try

                intMonth = CInt(HttpContext.Current.Request(strFrmFldName & "_month"))

                intDay = CInt(HttpContext.Current.Request(strFrmFldName & "_day"))

                intYear = CInt(HttpContext.Current.Request(strFrmFldName & "_year"))

                dtResult = CDate(intMonth & "/" & intDay & "/" & intYear)

                Return dtResult

            Catch ex As Exception

                Throw ex

            End Try

        End Function

        Public Shared Function EspMonth(ByVal intMonth As Integer) As String

            If intMonth > 0 And intMonth < 13 Then


                Dim arrMes() As String

                arrMes = Split("--,enero,febrero,marzo,abril,mayo,junio,julio,agosto,septiembre,octubre,noviembre,diciembre", ",")

                Return arrMes(intMonth)

            Else

                Return "Index out of range"

            End If

        End Function

        Public Shared Function DiplayArticleInfo(ByVal intPageId As Integer, ByVal strPageTitle As String, ByVal intPageRating As Integer, ByVal strPageAuthorName As String, ByVal datePosted As Date, ByVal strPageBlurb As String, Optional ByVal IsList As Boolean = True, Optional ByVal intIndex As Integer = 1, Optional ByVal strSearchParam As String = "") As String

            Dim sb As New StringBuilder("")

            Try

                If IsList Then sb.Append("<li>")

                sb.Append("<a href=" & _project_path & "article.aspx?page_id=" & intPageId & "&index=" & intIndex & "&search_param=" & strSearchParam & "><b>" & strPageTitle & "</b></a>")

                'This is the link if using AJAX
                'sb.Append("<a href=""javascript:"" onclick=""ShowAjaxContent('" & _project_path & "article.aspx?page_id=" & intPageId & "&index=" & intIndex & "&search_param=" & strSearchParam & "&format=print&close=y',530,200,this);"" >" & strPageTitle & "</a>")

                If intPageRating > 0 Then sb.Append("&nbsp;&nbsp;&nbsp;<img src=" & _project_path & "images/stars" & intPageRating & ".gif>")

                sb.Append("<br>")

                If strPageAuthorName <> "" Then sb.Append("por: <b>" & strPageAuthorName & "</b>")

                sb.Append(" (" & Functions.FormatHispanicDateTime(datePosted) & ")<br>")

                sb.Append(ReplaceCR(strPageBlurb))

                If IsList Then sb.Append("</li>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function ShowPicture(ByVal intUserId As Integer, ByVal strPictureName As String, Optional ByVal intSexoId As Integer = 0, Optional ByVal strClass As String = "") As String

            Dim sb As New StringBuilder("")

            'Dim strPic As String = ""
            'Dim strPic2 As String = ""
            Dim strFileName As String = "usr_" & intUserId & ".jpg"
            Dim strPath As String = HttpContext.Current.Request.PhysicalApplicationPath & "\files\"
            Dim strPathAlt As String = HttpContext.Current.Request.PhysicalApplicationPath & "\images\"

            Try

                'strPic = ProjectPath & "files/" & strPictureName

                'strPic2 = ProjectPath & "images/" & strPictureName



                'If System.IO.File.Exists(HttpContext.Current.Server.MapPath(strPic)) Then

                '    sb.Append(strPic)

                'ElseIf System.IO.File.Exists(HttpContext.Current.Server.MapPath(strPic2)) Then

                '    sb.Append(strPic2)

                'ElseIf System.IO.File.Exists(HttpContext.Current.Server.MapPath(strFileName)) Then



                'End If



                sb.Append("<img src=""")


                If System.IO.File.Exists(strPath & strFileName) Then

                    sb.Append(Functions.ProjectPath & "files/" & strFileName)

                ElseIf System.IO.File.Exists(strPath & strPictureName) Then

                    sb.Append(Functions.ProjectPath & "files/" & strPictureName)

                ElseIf System.IO.File.Exists(strPathAlt & strPictureName) Then

                    sb.Append(Functions.ProjectPath & "images/" & strPictureName)

                ElseIf intSexoId = 1 Then

                    sb.Append(Functions.ProjectPath & "images/AvatarMasculino.jpg")

                ElseIf intSexoId = 2 Then

                    sb.Append(Functions.ProjectPath & "images/AvatarFemenino.jpg")

                Else

                    sb.Append(Functions.ProjectPath & "images/AvatarMasculino.jpg")

                End If

                sb.Append("""") ' end of src="

                If strClass <> "" Then

                    sb.Append(" class=""" & strClass & """ ")

                End If

                sb.Append("/>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function DisplayAuthorInfo(ByVal intUserId As Integer, Optional ByVal dtPosted As Date = #1/1/1900#, Optional ByVal blShowPicture As Boolean = True) As String

            Dim sb As New System.Text.StringBuilder("")

            Dim strCacheVar As String = "Author" & intUserId

            Try

                If IsNothing(HttpContext.Current.Cache(strCacheVar)) Then

                    Dim objUser As DataLayer.reg_users = New DataLayer.reg_users(intUserId)

                    Try

                        If objUser.getFirstName <> "" Then

                            sb.Append("<p align=center>")

                            sb.Append("<a href=javascript:") 'Openwindow('" & _project_path & "bio.aspx?user_id=" & intUserId & "')>")

                            sb.Append(" onclick=""ShowAjaxContent('" & _project_path & "bio.aspx?user_id=" & intUserId & "&format=print&close=y',530,400,this);""")

                            sb.Append(">")

                            'If (objUser.getPicture <> "") And blShowPicture Then sb.Append("<img src=" & _project_path & "images/" & objUser.getPicture & " border=0><br>")

                            sb.Append(ShowPicture(intUserId, objUser.getPicture))

                            sb.Append("<br>")

                            sb.Append("Por: <i>" & objUser.getFirstName & " " & objUser.getLastName)

                        End If

                        HttpContext.Current.Cache.Insert(strCacheVar, sb.ToString())

                        sb.Remove(0, sb.Length)

                    Catch ex As Exception

                        ShowError(ex)

                    Finally

                        objUser = Nothing

                    End Try

                End If


                sb.Append(HttpContext.Current.Cache.Get(strCacheVar))

                If dtPosted <> #1/1/1900# Then sb.Append(" (" & FormatHispanicDateTime(dtPosted) & ")")

                sb.Append("</i></a>")

                sb.Append("</p>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function DisplayAuthorInfoHtml(ByVal intUserId As Integer, Optional ByVal dtPosted As Date = #1/1/1900#) As String

            Dim objUser As DataLayer.reg_users = New DataLayer.reg_users(intUserId)
            Dim sb As New System.Text.StringBuilder("")

            Try

                If objUser.getFirstName <> "" Then

                    sb.Append("<p align=center><a href=javascript:Openwindow('" & _project_path & "bio_" & intUserId & ".html')>")
                    If objUser.getPicture <> "" Then sb.Append("<img src=" & _project_path & "images/" & objUser.getPicture & " border=0><br>")
                    sb.Append("Por : <i>" & objUser.getFirstName & " " & objUser.getLastName)
                    If dtPosted <> #1/1/1900# Then sb.Append(" (" & FormatHispanicDateTime(dtPosted) & ")")
                    sb.Append("</i></a></p>")

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing
                objUser = Nothing

            End Try

        End Function

        Public Shared Function ShowFormPost() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim objContext As HttpContext = HttpContext.Current
            Dim objFld As Object

            Try

                For Each objFld In objContext.Request.Form()

                    sb.Append(objFld & ": " & objContext.Request(objFld) & "<br>")

                Next

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing
                objContext = Nothing

            End Try

        End Function

        Public Shared Function CheckBoxValToInt(ByVal strVal As String) As Integer

            If strVal = "on" Then

                Return 1

            Else

                Return 0

            End If

        End Function

        Public Shared Function CountArticles(ByVal strProcedure As String) As Integer

            Dim intTotalRecords As Integer = 0

            Try

                If IsNothing(HttpContext.Current.Cache(strProcedure)) Then

                    intTotalRecords = CInt(ParaLideres.GenericDataHandler.ExecScalar(strProcedure))

                    HttpContext.Current.Cache.Insert(strProcedure, intTotalRecords)

                Else

                    intTotalRecords = CInt(HttpContext.Current.Cache.Get(strProcedure))

                End If

                Return intTotalRecords

            Catch ex As Exception

                Throw ex

            End Try

        End Function


        Public Shared Function ShowArticles(ByVal strProcedure As String, ByVal intTotalRecords As Integer, ByVal intIndex As Integer, Optional ByVal strFieldName As String = "", Optional ByVal strFieldValue As String = "") As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                If intTotalRecords > 0 Then

                    Dim reader As System.Data.SqlClient.SqlDataReader

                    Dim intPageId As Integer = 0
                    Dim intPageRating As Integer = 0

                    Dim datePosted As Date = #1/1/1900#

                    Dim strPageTitle As String = ""
                    Dim strPageBlurb As String = ""
                    Dim strPageAuthorName As String = ""

                    Try

                        reader = ParaLideres.GenericDataHandler.GetRecords(strProcedure)

                        If reader.HasRows Then

                            sb.Append("<ul class=UL>")

                            Do While (reader.Read())

                                intPageRating = 0

                                If Not reader.IsDBNull(1) Then intPageId = reader(1) Else intPageId = 0
                                If Not reader.IsDBNull(2) Then datePosted = reader(2) Else datePosted = #1/1/1900#
                                If Not reader.IsDBNull(3) Then strPageTitle = reader(3) Else strPageTitle = ""
                                If Not reader.IsDBNull(4) Then strPageBlurb = Functions.ReplaceCR(reader(4)) Else strPageBlurb = ""
                                If Not reader.IsDBNull(5) Then intPageRating = reader(5) Else intPageRating = 0
                                If Not reader.IsDBNull(6) Then strPageAuthorName = reader(6) Else strPageAuthorName = ""

                                sb.Append(Functions.DiplayArticleInfo(intPageId, strPageTitle, intPageRating, strPageAuthorName, datePosted, strPageBlurb, True, intIndex))

                            Loop

                            sb.Append("</ul>")

                        End If

                        sb.Append("<p align=center>" & BottomMenu(intTotalRecords, intIndex, strFieldName, strFieldValue) & "</p>")

                    Catch ex As Exception

                        Throw ex

                    Finally

                        reader.Close()
                        reader = Nothing

                    End Try

                Else

                    sb.Append("No encontramos art&#237;culos relacionados a tu b&#250;squeda.  Por favor no uses art&#237;culos (el, la, los, un, uno, etc.) ni tampoco uses palabras como 'AND' o 'OR'")

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function ShowAllArticles(ByVal strProcedure As String) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                Dim reader As System.Data.SqlClient.SqlDataReader

                Dim intPageId As Integer
                Dim datePosted As Date
                Dim strPageTitle As String
                Dim strPageBlurb As String
                Dim intPageRating As Integer
                Dim strPageAuthorName As String

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords(strProcedure)

                    If reader.HasRows Then

                        sb.Append("<ul class=UL>")

                        Do While (reader.Read())

                            intPageRating = 0

                            If Not reader.IsDBNull(1) Then intPageId = reader(1)
                            If Not reader.IsDBNull(2) Then datePosted = reader(2)
                            If Not reader.IsDBNull(3) Then strPageTitle = reader(3)
                            If Not reader.IsDBNull(4) Then strPageBlurb = Functions.ReplaceCR(reader(4))
                            If Not reader.IsDBNull(5) Then intPageRating = reader(5)
                            If Not reader.IsDBNull(6) Then strPageAuthorName = reader(6)

                            sb.Append(Functions.DiplayArticleInfo(intPageId, strPageTitle, intPageRating, strPageAuthorName, datePosted, strPageBlurb))

                        Loop

                        sb.Append("</ul>")

                    Else

                        sb.Append("No encontramos art&#237;culos relacionados a tu b&#250;squeda.")

                    End If

                Catch ex As Exception

                    ShowError(ex)

                Finally

                    reader.Close()
                    reader = Nothing

                End Try

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function LoDestacado(ByVal intIndex As Integer) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                strProcedure = "proc_GetLoDestacado " & intIndex
                intCount = CountArticles("proc_LoDestacadoCount")

                Return ShowArticles(strProcedure, intCount, intIndex)

            Catch ex As Exception

                ShowError(ex)

            End Try



        End Function

        Public Shared Function LoUltimo(ByVal intIndex As Integer) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                strProcedure = "proc_GetLoUltimo " & intIndex
                intCount = CountArticles("proc_LoUltimoCount")

                Return ShowArticles(strProcedure, intCount, intIndex)

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function

        Public Shared Function PagesByTag(ByVal intIndex As Integer, ByVal strTag As String) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                strProcedure = "proc_GetPagesByTags " & intIndex & ",'" & strTag & "'"
                intCount = CountArticles("proc_GetPagesByTagsCount '" & strTag & "'")

                Return ShowArticles(strProcedure, intCount, intIndex)

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function


        Public Shared Function MisFavoritos(ByVal intIUserId As Integer, ByVal intIndex As Integer) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                strProcedure = "proc_GetMyFavorites " & intIUserId & "," & intIndex
                intCount = CountArticles("proc_MyFavoritesCount " & intIUserId)

                Return ShowArticles(strProcedure, intCount, intIndex)

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function


        Public Shared Function GetAllSubSections(ByVal intSectionId As Integer, Optional ByVal blRemoveLast As Boolean = True) As String

            Dim sb As New System.Text.StringBuilder("")

            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intSubSectionId As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsByParentId " & intSectionId)

                If reader.HasRows Then

                    Do While reader.Read()

                        If Not reader.IsDBNull(0) Then

                            intSubSectionId = reader(0)

                            sb.Append(intSubSectionId & ",")

                            sb.Append(GetAllSubSections(intSubSectionId, False))

                        End If

                    Loop

                    If blRemoveLast Then sb.Remove(sb.Length - 1, 1)

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try
        End Function

        Public Shared Function IsPalindrome(ByVal strVal As String) As Boolean

            Dim sb As System.Text.StringBuilder

            Dim blIsPal As Boolean = True

            Dim intLen As Integer = 0
            Dim intIndex As Integer = 0
            Dim intTop As Integer = 0
            Dim intMirror As Integer = 0

            Dim charVar As Char

            Try

                sb = New System.Text.StringBuilder("")

                If strVal.Length() > 1 Then

                    For intIndex = 1 To strVal.Length

                        charVar = Mid(strVal, intIndex, 1)

                        If (Asc(charVar) >= 97 And Asc(charVar) <= 122) Or (Asc(charVar) >= 65 And Asc(charVar) <= 90) Then

                            sb.Append(charVar)

                        End If

                    Next

                    strVal = LCase(sb.ToString())

                    HttpContext.Current.Trace.Write("sb: " & sb.ToString())

                    intLen = strVal.Length()

                    intTop = CInt(intLen / 2) - 1

                    HttpContext.Current.Trace.Write("intLen: " & intLen)
                    HttpContext.Current.Trace.Write("intTop: " & intTop)

                    For intIndex = 0 To intTop

                        intMirror = intLen - intIndex - 1

                        HttpContext.Current.Trace.Write("intIndex: " & intIndex)
                        HttpContext.Current.Trace.Write("intMirror: " & intMirror)


                        HttpContext.Current.Trace.Write("strVal.Chars(intIndex): " & strVal.Chars(intIndex))
                        HttpContext.Current.Trace.Write("strVal.Chars(intMirror): " & strVal.Chars(intMirror))

                        If strVal.Chars(intIndex) <> strVal.Chars(intMirror) Then

                            blIsPal = False

                        End If

                    Next

                End If

                Return blIsPal

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function ClearString(ByVal strVal As String) As String

            HttpContext.Current.Trace.Write("Original strVal: " & strVal)

            Dim sb As System.Text.StringBuilder

            Dim intLen As Integer = 0
            Dim intIndex As Integer = 0

            Dim charVar As Char

            Try

                sb = New System.Text.StringBuilder("")

                If strVal.Length() > 1 Then

                    For intIndex = 1 To strVal.Length

                        charVar = Mid(strVal, intIndex, 1)

                        If (Asc(charVar) >= 97 And Asc(charVar) <= 122) Or (Asc(charVar) >= 65 And Asc(charVar) <= 90) Or (Asc(charVar) = 32) Then

                            sb.Append(charVar)

                        End If

                    Next

                    strVal = sb.ToString()

                    HttpContext.Current.Trace.Write("Final strVal: " & strVal)

                End If

                Return strVal

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Shared Function SearchResults(ByVal strParam As String, ByVal intIndex As Integer) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                'strParam = ClearString(Trim(strParam))

                'NOTE: Use '"" around search param so that two words can be used together when using contains

                strProcedure = "proc_Search '" & strParam & "'," & intIndex

                intCount = CountArticles("proc_SearchCount '" & strParam & "'")

                HttpContext.Current.Trace.Write("strProcedure :" & strProcedure)

                HttpContext.Current.Trace.Write("intCount :" & intCount)

                Return ShowArticles(strProcedure, intCount, intIndex, "shearchparam", strParam)

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function
        Public Shared Function ArticlesByUser(ByVal intUserId As Integer, ByVal intIndex As Integer) As String

            Dim strProcedure As String = ""
            Dim intCount As Integer = 0

            Try

                strProcedure = "proc_GetArtiblesByUser " & intUserId & "," & intIndex
                intCount = CountArticles("proc_ArticlesByUserCount " & intUserId)

                Return ShowArticles(strProcedure, intCount, intIndex, "user_id", intUserId)

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function

        Public Shared Function camelNotation(ByVal strVal As String, ByVal fromStart As Boolean) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim intLen As Integer = 0
            Dim intStartIndex As Integer = 1
            Dim X As Integer = 0
            Dim thisCar As Char

            Try

                intLen = Len(strVal)

                If fromStart Then

                    sb.Append(UCase(Left(strVal, 1)))

                Else

                    sb.Append(LCase(Left(strVal, 1)))

                End If

                intStartIndex = 2

                For X = intStartIndex To intLen

                    thisCar = Mid(strVal, X, 1)

                    If thisCar = "_" Then

                        X = X + 1
                        thisCar = Mid(strVal, X, 1)
                        sb.Append(UCase(thisCar))

                    Else

                        sb.Append(LCase(thisCar))

                    End If

                Next

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function


        Public Shared Function CreateFormLabel(ByVal strVal As String) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim intLen As Integer = 0
            Dim intStartIndex As Integer = 1
            Dim X As Integer = 0
            Dim thisCar As Char

            Try

                intLen = Len(strVal)

                sb.Append(UCase(Left(strVal, 1)))

                intStartIndex = intStartIndex + 1

                For X = intStartIndex To intLen

                    thisCar = Mid(strVal, X, 1)

                    If thisCar = "_" Then

                        sb.Append(" ")

                        X = X + 1
                        thisCar = Mid(strVal, X, 1)
                        sb.Append(UCase(thisCar))

                    Else

                        sb.Append(LCase(thisCar))

                    End If

                Next

                sb.Replace("_", " ")

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function


        Public Shared Function BottomMenu(ByVal intRecords As Integer, ByVal intIndex As Integer, Optional ByVal strFieldName As String = "", Optional ByVal strFieldValue As String = "") As String

            Dim intPages As Integer = 0
            Dim intPageSize As Integer = 10
            Dim intCurrentPage As Integer = 0
            Dim _strPrev As String = ""
            Dim _strNext As String = ""
            Dim strNumericMenu As String = ""
            Dim intX As Integer = 0
            Dim intMenuStart As Integer = 0
            Dim intMenuEnd As Integer = 0
            Dim strPageUrl As String = HttpContext.Current.Request.Path()
            Dim strExtraLink As String = ""

            If intMenuEnd > (intPages - 1) Then intMenuEnd = (intPages - 1)


            If strFieldName <> "" And strFieldValue <> "" Then

                strExtraLink = "&" & strFieldName & "=" & strFieldValue

            End If


            intPages = Math.Ceiling(intRecords / intPageSize)

            If intIndex > intRecords Then

                intIndex = intRecords - intPageSize

            End If

            intCurrentPage = CInt((intIndex + intPageSize - 1) / intPageSize)

            intMenuStart = (Int((intCurrentPage - 1) / 10) * 10)

            intMenuEnd = intMenuStart + 9

            If intMenuEnd > (intPages - 1) Then intMenuEnd = (intPages - 1)

            For intX = intMenuStart To intMenuEnd

                If intCurrentPage = intX + 1 Then

                    strNumericMenu = strNumericMenu & " <b>" & intX + 1 & "</b> "

                Else

                    strNumericMenu = strNumericMenu & " <a href=""" & strPageUrl & "?index=" & (intX * intPageSize) + 1 & strExtraLink & """>" & intX + 1 & "</a> "

                End If

            Next

            If intPages > 1 Then

                Select Case intCurrentPage

                    Case Is <= 1

                        _strPrev = "<< Previo " & strNumericMenu
                        _strNext = "<a href=""" & strPageUrl & "?index=" & intIndex + intPageSize & strExtraLink & """>Siguiente >></a>"

                    Case Is >= intPages

                        _strPrev = "<a href=""" & strPageUrl & "?index=" & intIndex - intPageSize & strExtraLink & """><< Previo</a> " & strNumericMenu
                        _strNext = "Siguiente >>"

                    Case Else

                        _strPrev = "<a href=""" & strPageUrl & "?index=" & intIndex - intPageSize & strExtraLink & """><< Previo</a> " & strNumericMenu
                        _strNext = "<a href=""" & strPageUrl & "?index=" & intIndex + intPageSize & strExtraLink & """>Siguiente >></a>"

                End Select

            End If

            Return _strPrev & " " & _strNext

        End Function

        Public Shared Function DisplayArticleRating(ByVal intPageId As Integer, ByVal intSectionId As Integer, ByVal strSectionName As String, ByVal intIndex As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intStarNum As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("sp_GetRatingByPageID " & intPageId)

                If reader.HasRows Then

                    reader.Read()

                    If Not reader.IsDBNull(0) Then intStarNum = reader(0)

                    'sb.Append("<p align=center>| <a href=section.aspx?section_id=" & intSectionId & "&index=" & intIndex & ">Regresar a '" & strSectionName & "'</a> | </p>")

                    sb.Append("<p align=center>| ")

                    If intStarNum > 0 Then sb.Append("Rating: <img src=" & _project_path & "images/stars" & intStarNum & ".gif border=0>  |  ")

                    sb.Append("<a href=" & _project_path & "article.aspx?page_id=" & intPageId & "&format=print target=new>Imprimir P&#225;gina <img src=" & _project_path & "images/print.gif border=0 align=bottom></a> | ")

                    sb.Append("</p>")

                    sb.Append("<br><img src=""" & _project_path & "_images/puntos_horizontales.gif"" >")

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                reader.Close()
                reader = Nothing
                sb = Nothing

            End Try



        End Function

        Public Shared Function UpdateTotalNumberOfArticlesForAllSite() As String

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As New System.Text.StringBuilder("")

            Dim intSectionId As Integer = 0
            Dim strSectionName As String = ""
            Dim intTotal As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("Select section_id, section_name from sections order by section_name")

                If reader.HasRows() Then

                    Do While (reader.Read())

                        intTotal = 0

                        intSectionId = reader(0)
                        strSectionName = reader(1)

                        Try
                            intTotal = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_CountPostedArticlesBySectionId " & intSectionId))
                        Catch ex As Exception
                        End Try

                        intTotal = intTotal + GetTotalNoOfArticles(intSectionId)

                        UpdateTotalNumberOfArticlesForSection(intSectionId, intTotal)

                        sb.Append("<br>" & strSectionName & " was updated with " & intTotal)

                    Loop

                End If

                Return sb.ToString


            Catch ex As Exception

                Throw ex

            Finally

                reader.Close()
                reader = Nothing

                sb = Nothing

            End Try



        End Function

        Public Shared Sub UpdateTotalNumberOfArticlesForSection(ByVal intSectionId As Integer, ByVal intTotal As Integer)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateTotalNoOfArticles " & intSectionId & "," & intTotal)

            Catch ex As Exception

                Throw ex

            End Try


        End Sub

        Public Shared Sub UpdateTotalNumberOfArticlesForSection(ByVal intSectionId As Integer)

            Dim intTotal As Integer = 0

            Try

                Try
                    intTotal = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_CountPostedArticlesBySectionId " & intSectionId))
                Catch ex As Exception
                End Try

                intTotal = intTotal + GetTotalNoOfArticles(intSectionId)

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateTotalNoOfArticles " & intSectionId & "," & intTotal)

            Catch ex As Exception

                Throw ex

            End Try


        End Sub


        Public Shared Function GetTotalNoOfArticles(ByVal intSectionId As Integer) As Integer

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intTotal As Integer = 0
            Dim intThisSection As Integer = 0
            Dim intThisTotal As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_CountPagesBySectionId " & intSectionId)

                If reader.HasRows Then

                    Do While (reader.Read())

                        intThisSection = reader(0)
                        intThisTotal = reader(1)

                        intTotal = intTotal + intThisTotal + GetTotalNoOfArticles(intThisSection)

                    Loop

                End If

                Return intTotal

            Catch ex As Exception

                Throw ex

            Finally

                reader.Close()
                reader = Nothing

            End Try

        End Function


#Region "Formatting"

        Public Shared Sub ShowError(ByVal excError As Exception, Optional ByVal strMethod As String = "")

            Dim sb As New System.Text.StringBuilder("")

            Dim objErr As ParaLideres.ErrorHandler

            Try

                objErr = New ParaLideres.ErrorHandler

                sb.Append("<br>" & strMethod)

                Try

                    sb.Append("<br>" & objErr.ReturnHtmlErrorMessage(excError))

                Catch ex As Exception

                End Try

                If System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode") Then

                    HttpContext.Current.Response.Write(sb.ToString())

                Else

                    SendMail("apoyo@paralideres.org", _develop, "Error on Paralideres - " & strMethod, sb.ToString())

                End If

            Catch ex As Exception

            Finally

                sb = Nothing

                objErr = Nothing

            End Try

        End Sub


        Public Shared Function ReplaceCR(ByVal ps_string As String) As String

            'ps_string = HttpContext.Current.Server.HtmlDecode(ps_string)
            ps_string = Replace(ps_string, Chr(13), "<br>")

            Return ps_string

        End Function

        Public Shared Function FormatString(ByVal ps_string As String) As String

            If ps_string <> "" Then

                'ps_string = HttpContext.Current.Server.HtmlEncode(ps_string)
                ps_string = HttpContext.Current.Server.HtmlDecode(ps_string)
                ps_string = Replace(ps_string, Chr(34), "&#34;")
                ps_string = Replace(ps_string, Chr(39), "&#39;")

                'ps_string = Replace(ps_string, "&#225;", "á")
                'ps_string = Replace(ps_string, "&#233;", "é")
                'ps_string = Replace(ps_string, "&#237;", "í")
                'ps_string = Replace(ps_string, "&#243;", "ó")
                'ps_string = Replace(ps_string, "&#250;", "ú")

                ps_string = Replace(ps_string, "á", "&#225;")
                ps_string = Replace(ps_string, "é", "&#233;")
                ps_string = Replace(ps_string, "í", "&#237;")
                ps_string = Replace(ps_string, "ó", "&#243;")
                ps_string = Replace(ps_string, "ú", "&#250;")

                ps_string = Trim(ps_string)

            Else

                ps_string = " "

            End If

            Return ps_string

        End Function

        Public Shared Function FormatCase(ByVal strVar As String, Optional ByVal AllFields As Boolean = False) As String

            Dim sb As New StringBuilder("")
            Dim intIndex As Integer = 0
            Dim arrWords As String()

            Try

                If AllFields Then

                    arrWords = Split(strVar, Chr(32))

                    For intIndex = 0 To UBound(arrWords)

                        sb.Append(UCase(Left(arrWords(intIndex), 1)))

                        sb.Append(LCase(Mid(arrWords(intIndex), 2)))

                        If intIndex < UBound(arrWords) Then sb.Append(Chr(32))

                    Next

                Else

                    sb.Append(UCase(Left(strVar, 1)))

                    sb.Append(LCase(Mid(strVar, 2)))

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

#End Region

#Region "Mail"


        Public Shared Function SendMail2(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal message As String, Optional ByVal strFiles As String = "") As Boolean

            Dim Mail As Object

            Try

                Mail = HttpContext.Current.Server.CreateObject("Persits.MailSender")

                Mail.Host = "mail.yourdomain.com"

                'Mail.Username = "emailaddress@yourdomain.com"
                'Mail.Password = "your password"

                Mail.From = "webmaster@yourdomain.com"
                Mail.FromName = "Your Name"
                Mail.AddAddress("recipient@somedomain.com")
                Mail.AddCC("you@yourdomain.com")
                Mail.Subject = "Subject goes here"
                Mail.Body = "Message body goes here"

                Mail.Queue = 1

                'Mail.Send(Missing.Value)

            Catch ex As Exception

            End Try



        End Function


        Public Shared Function SendMail(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal message As String, Optional ByVal strFiles As String = "") As Boolean

            Dim sb As New System.Text.StringBuilder("")
            Dim objClient As System.Net.Mail.SmtpClient
            Dim objMessage As New System.Net.Mail.MailMessage
            Dim objAddressFrom As System.Net.Mail.MailAddress
            Dim objAddressTo As System.Net.Mail.MailAddress

            Dim arrFiles() As String

            Dim intIndex As Integer

            Dim blResults As Boolean = False

            Dim strUrl As String = ""


            Try

                objClient = New System.Net.Mail.SmtpClient(System.Web.Configuration.WebConfigurationManager.AppSettings("MailServer"))

                strUrl = "http://" & HttpContext.Current.Request.ServerVariables("HTTP_HOST")

                HttpContext.Current.Trace.Write("BASE URL: " & strUrl)

                message = message.Replace("src=/", "src=" & strUrl & "/")
                message = message.Replace("src=""/", "src=""" & strUrl & "/")
                message = message.Replace("src='/", "src='" & strUrl & "/")

                message = message.Replace("href=/", "href=" & strUrl & "/")
                message = message.Replace("href=""/", "href=""" & strUrl & "/")
                message = message.Replace("href='/", "href='" & strUrl & "/")

                sb.Append("<html><head><title>" & subject & "</title>" & vbLf)

                sb.Append("<script language=""javascript"" src=""" & strUrl & _project_path & "ajax.js"" type=""text/javascript""></script>" & vbLf)

                sb.Append("<link rel=""stylesheet"" href=""" & strUrl & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""stylesheet"" href=""" & strUrl & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)

                sb.Append("</head>" & Chr(13))

                sb.Append("<body topmargin=""0"" leftmargin=""0"">" & Chr(13))

                sb.Append("<p align=""center""><a href=" & strUrl & _project_path & "home.aspx><img src=""" & strUrl & _project_path & "_images/header_contenido.jpg"" width=""530"" height=""67"" border=0></a></p>")

                sb.Append("<p align=""left"">")

                sb.Append(Replace(message, Chr(13), "<br>"))

                sb.Append("</p>")

                sb.Append("</body></html>")



                'objClient.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials

                objAddressFrom = New System.Net.Mail.MailAddress(sender)
                objAddressTo = New System.Net.Mail.MailAddress(recipient)

                objMessage.Priority = Net.Mail.MailPriority.Normal
                objMessage.From = objAddressFrom
                objMessage.Subject = subject
                objMessage.Body = sb.ToString()
                objMessage.IsBodyHtml = True

                If System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode") Then

                    objMessage.To.Add(System.Web.Configuration.WebConfigurationManager.AppSettings("SupportAccount"))

                Else

                    objMessage.To.Add(objAddressTo)

                End If

                If strFiles <> "" Then

                    arrFiles = Split(strFiles, ",")

                    For intIndex = 0 To UBound(arrFiles)

                        If System.IO.File.Exists(arrFiles(intIndex)) Then

                            objMessage.Attachments.Add(New System.Net.Mail.Attachment(arrFiles(intIndex)))

                        End If

                    Next

                End If

                Try

                    HttpContext.Current.Trace.Write("FROM: " & sender)
                    HttpContext.Current.Trace.Write("TO: " & recipient)
                    HttpContext.Current.Trace.Write("SUBJECT: " & subject)
                    HttpContext.Current.Trace.Write("MESSAGE: " & sb.ToString())
                    HttpContext.Current.Trace.Write("HOST: " & objClient.Host)

                    objClient.Send(objMessage)

                    blResults = True

                    HttpContext.Current.Trace.Write("E-mail was sent successfully!")

                Catch ex As Exception

                    HttpContext.Current.Trace.Warn(ex.ToString())

                End Try

                Return blResults

            Catch ex As Exception

                HttpContext.Current.Trace.Warn(ex.ToString())

                'ShowError(ex)

            Finally

                objMessage = Nothing

                objClient = Nothing

                objAddressFrom = Nothing

                objAddressTo = Nothing

            End Try

        End Function

        Public Shared Function MassMailing(ByVal strStoredProc As String, ByVal strSubject As String, ByVal strMessage As String, Optional ByVal isTest As Boolean = False) As String

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As New StringBuilder("")
            Dim strRecipient = ""

            reader = ParaLideres.GenericDataHandler.GetRecords(strStoredProc)

            If reader.HasRows Then

                Do While reader.Read()

                    strRecipient = reader.GetString(0)

                    If Not isTest Then 'if isTest = False then send e-mail

                        If SendMail(_support_account, strRecipient, strSubject, strMessage) Then

                            sb.Append("Message successfully sent to: " & strRecipient & "<br>")

                        Else 'if isTest = True then do not send e-mail

                            sb.Append("Error: failed to send message to: " & strRecipient & "<br>")

                        End If

                    Else

                        sb.Append("Message would have been sent to: " & strRecipient & "<br>")

                    End If

                Loop

                'Always send a copy of the message to Xavier
                If SendMail(_support_account, "xc2k@hotmail.com", strSubject, strMessage) Then

                    sb.Append("<p>Message was sent to Xavier")

                Else

                    sb.Append("<p>Message was NOT sent to Xavier")

                End If

                'Display message sent
                sb.Append("<p>")
                sb.Append("<b><u>Message That Was Sent</u></b><br>")
                sb.Append("<b>Subject: </b>" & strSubject & "<br>")
                sb.Append("<b>Message: </b>" & strMessage & "<br>")

            End If


            reader = Nothing

            Return sb.ToString()

        End Function

#End Region

#Region "Database"

        Public Shared Function DisplayDBTable(ByVal strSQL As String, Optional ByVal strConfirmName As String = "", Optional ByVal strUrl As String = "", Optional ByVal strTableHeader As String = "", Optional ByVal boolShowMessage As Boolean = False, Optional ByVal strAlign As String = "left", Optional ByVal intWidth As Integer = 500, Optional ByVal strActionToTake As String = "delete") As String


            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
            Dim intIndex As Integer = 0
            Dim strBgColor As String = "white"
            Dim strScript As String = ""

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords(strSQL)

                If reader.HasRows Then

                    sb.Append("<table cellpadding=2 cellspacing=0 rules=""rows"" bordercolor=""Black"" border=""1"" style=""border-color:Black;border-width:1px;border-style:solid;width:" & intWidth & "px;border-collapse:collapse;"">")
                    'sb.Append(SetTableProperties(intWidth))

                    If strTableHeader <> "" Then

                        sb.Append("<tr class=WHITEBOLD bgcolor=" & _strBackGroundColor & ">")
                        sb.Append("<td colspan=" & reader.FieldCount() & " align=center>" & UCase(strTableHeader) & "</td>")
                        sb.Append("</tr>")

                    End If

                    'generate header
                    sb.Append("<tr class=WHITEBOLD bgcolor=" & _strBackGroundColor & ">")

                    For intIndex = 0 To reader.FieldCount() - 1

                        sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                    Next

                    sb.Append("</tr>")


                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        sb.Append("<tr class=GEN  " & strScript & ">")

                        For intIndex = 0 To reader.FieldCount() - 1

                            HttpContext.Current.Trace.Write(reader.GetName(intIndex) & ": " & reader(intIndex).GetType.ToString)

                            If Not IsDBNull(reader(intIndex)) Then

                                Select Case reader(intIndex).GetType.ToString

                                    Case "System.Double"

                                        sb.Append("<td valign=top align=right>" & FormatNumber(reader(intIndex), 2) & "</td>")

                                    Case "System.DateTime"

                                        sb.Append("<td valign=top align=center>" & FormatHispanicDateTime(reader(intIndex)) & "</td>")

                                    Case Else

                                        sb.Append("<td valign=top align=" & strAlign & ">" & HttpContext.Current.Server.HtmlDecode(reader(intIndex)) & "</td>")

                                End Select

                            Else

                                sb.Append("<td valign=top align=" & strAlign & ">N/A</td>")

                            End If

                        Next

                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBackGroundColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")


                    'If strConfirmName <> "" And strUrl <> "" Then

                    '    sb.Append(ConfirmDelete(strConfirmName, strUrl, strActionToTake))

                    'End If

                Else

                    If boolShowMessage Then sb.Append("There are no records that matched your query. Please try again.")

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                reader.Close()
                reader = Nothing
                sb = Nothing

            End Try

        End Function

#End Region

        Public Shared Function GetRandomNumber(ByVal intLowerBound As Integer, ByVal intUpperBound As Integer) As Integer

            ' Returns a random number between intLowerBound and intUpperBound

            Dim objRandom As New System.Random

            Try

                objRandom = New System.Random

                Return objRandom.Next(intLowerBound, intUpperBound)

            Catch ex As Exception

                ShowError(ex)

            Finally

                objRandom = Nothing

            End Try

        End Function

        Public Shared Function GenerateCell(ByVal strContent As String, Optional ByVal intColType As ColType = ColType.Text, Optional ByVal intColSpan As Integer = 1, Optional ByVal strClass As String = "GEN", Optional ByVal strAlign As String = "left", Optional ByVal strWrap As String = "", Optional ByVal strBgColor As String = "") As String

            Dim sb As New System.Text.StringBuilder("")
            Dim strFormat As String = " x:str "

            Try

                Select Case intColType

                    Case ColType.Number

                        strContent = FormatNumber(strContent, 2)
                        strAlign = "right"
                        strFormat = "x:num"
                        strWrap = " nowrap "

                    Case ColType.Money

                        strContent = "$" & FormatNumber(strContent, 2)
                        strAlign = "right"
                        strFormat = "x:num"
                        strWrap = " nowrap "

                    Case ColType.Percentage

                        strContent = FormatNumber(strContent, 2) & "%"
                        strAlign = "right"
                        strFormat = "x:num"
                        strWrap = " nowrap "

                    Case ColType.DateType

                        strContent = FormatHispanicDateTime(strContent)
                        strAlign = "center"
                        strWrap = " nowrap "

                    Case ColType.TextAlignCenter

                        strAlign = "center"

                    Case ColType.TextAlignRight

                        strAlign = "right"

                    Case ColType.ImageCheck

                        strAlign = "center"


                        If strContent = 1 Then

                            strContent = "<img src=" & ProjectPath & "images/post_yes.gif>"

                        Else

                            strContent = "<img src=" & ProjectPath & "images/post_no.gif>"

                        End If

                End Select

                sb.Append("<td class=" & strClass & " align=" & strAlign & " valign=bottom " & strFormat)

                If intColSpan > 1 Then sb.Append(" colspan=" & intColSpan & " ")

                sb.Append(strWrap)

                If strBgColor <> "" Then sb.Append(" bgcolor=" & strBgColor)

                sb.Append(" >")

                sb.Append(strContent)

                sb.Append("</td>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Shared Function ShowFeatured(ByVal intListId As Integer) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intSectionId As Integer = 0
            Dim intPageId As Integer = 0
            Dim intAuthorId As Integer = 0
            Dim intType As Integer = 0

            Dim strSectionName As String = ""
            Dim strPageTitle As String = ""
            Dim strAuthor As String = ""
            Dim strLink As String = ""

            Dim dtPosted As Date = #1/1/0900#

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetFeatured " & intListId)

                If reader.HasRows() Then

                    Do While reader.Read()

                        intPageId = reader(0)
                        intSectionId = reader(1)
                        strSectionName = reader(2)
                        strPageTitle = reader(3)
                        intType = reader(4)
                        intAuthorId = reader(5)
                        strAuthor = reader(6)
                        dtPosted = reader(7)

                        sb.Append("                           <div class=""lu_content_box"">" & vbCrLf)

                        strLink = ParaLideres.Functions.ProjectPath & "section.aspx?section_id=" & intSectionId

                        sb.Append("                            	<a href=""" & strLink & """ class=""link_btn_round_sm"">" & strSectionName & "</a>" & vbCrLf)

                        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & ProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

                        strLink = ParaLideres.Functions.ProjectPath & "article.aspx?page_id=" & intPageId

                        sb.Append("                                <h1><a href=""" & strLink & """>" & strPageTitle & "</a></h1>" & vbCrLf)

                        strLink = ParaLideres.Functions.ProjectPath & "bio.aspx?user_id=" & intAuthorId

                        sb.Append("                                <span>Por <a href=""" & strLink & """>" & strAuthor & "</a> el " & ParaLideres.Functions.FormatHispanicDateTime(dtPosted) & "</span>" & vbCrLf)

                        sb.Append("                            </div>" & vbCrLf)

                    Loop

                    sb.Append("                    <div class=""tabs_more"">" & vbCrLf)

                    Select Case intListId

                        Case 1

                            strLink = "lo_ultimo.aspx"

                        Case 2

                            strLink = "destacado.aspx"

                        Case 3

                            strLink = ""


                    End Select

                    sb.Append("                    	<a href=""" & ParaLideres.Functions.ProjectPath & strLink & """>Ver Más</a>" & vbCrLf)

                    sb.Append("                    </div>" & vbCrLf)

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

                reader.Close()
                reader = Nothing

            End Try

        End Function

        Public Shared Function MenuDisplayBox(ByVal intSectionId As Integer) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim reader2 As System.Data.SqlClient.SqlDataReader


            Dim intSubsectionId As Integer = 0
            Dim intChildId As Integer = 0

            Dim strSubsectionName As String = ""
            Dim strChildName As String = ""
            Dim strPath As String = ""

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsByParentId " & intSectionId)

                sb.Append("             <div style=""text-align:right;""><a href=""javascript:ClearAjax('menu_display_box');""><img src=""" & Functions.ProjectPath & "images/x.gif"" border=""0"" vspace=""5"" hspace=""5"" alt=""Cerrar Menú"" title=""Cerrar Menú""></a></div>" & vbCrLf)

                If reader.HasRows() Then

                    sb.Append("            <ul class=""mdb_list"">" & vbCrLf)

                    Do While reader.Read()

                        intSubsectionId = reader(0)
                        strSubsectionName = reader(1)

                        strPath = Functions.ProjectPath & "section.aspx?section_id=" & intSubsectionId

                        sb.Append("            	<li>" & vbCrLf)

                        sb.Append("                	<ul>" & vbCrLf)

                        sb.Append("      	            <li class=""mdb_sec_title""><a href=""" & strPath & """>" & strSubsectionName & "</a></li>" & vbCrLf)

                        reader2 = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsByParentId " & intSubsectionId)

                        If reader2.HasRows() Then

                            Do While reader2.Read()

                                intChildId = reader2(0)
                                strChildName = reader2(1)

                                strPath = Functions.ProjectPath & "section.aspx?section_id=" & intChildId

                                sb.Append("                    	  <li><a href=""" & strPath & """>" & strChildName & "</a></li>" & vbCrLf)

                            Loop

                        End If

                        sb.Append("                    </ul>" & vbCrLf)

                        sb.Append("         </li>" & vbCrLf)

                    Loop

                    sb.Append("            </ul>" & vbCrLf)

                End If

                Return sb.ToString()

            Catch ex As Exception

                HttpContext.Current.Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

                reader.Close()
                reader = Nothing

                reader2.Close()
                reader2 = Nothing

            End Try

        End Function


    End Class




End Namespace
