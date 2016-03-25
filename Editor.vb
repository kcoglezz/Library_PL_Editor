Imports System
Imports System.Web
Imports System.Web.Mail
Imports System.Text
'Imports ParaLideres.Functions
Imports System.Web.UI
Imports Library_PL_Editor.ParaLideres.Functions

Namespace ParaLideres

    Public Class Editor

        Inherits Page

#Region "Declarations"


        Public _objUser As DataLayer.reg_users

        Public _blIsDebugMode As String = System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode")
        Public _strAlpha As String = ""

        Private _strPageContent As String = ""
        Private _strPageTitle As String = ""
        Private _strProjectPath1 As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _strProjectPath As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath_editor")
        Private _strBgColor As String = "#F9F1B1" '"#F9E07C"

        Public qs As SecureQueryString

        Public Enum Rules As Integer

            None = 0
            Rows = 1
            Cols = 2
            All = 3

        End Enum

        Private Enum FldType As Integer

            TextBox = 1
            TextArea = 2
            DropDownList = 3
            File = 4
            Calendar = 5

        End Enum

#End Region

#Region "Properties"

        Public WriteOnly Property PageContent() As String

            Set(ByVal Value As String)

                _strPageContent = Value

            End Set

        End Property

        Public WriteOnly Property PageTitle() As String

            Set(ByVal Value As String)

                _strPageTitle = Value

            End Set

        End Property

        Public WriteOnly Property BackColor() As String

            Set(ByVal Value As String)

                _strBgColor = Value

            End Set

        End Property

        Public ReadOnly Property ProjectPath() As String
            Get
                Return _strProjectPath
            End Get
        End Property

#End Region

#Region "Functions"

        Public Function DisplayReport(ByVal strSql As String, ByVal isSortable As Boolean, Optional ByVal strSortBy As String = "", Optional ByVal strDirection As String = "ASC") As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intColCount As Integer = 0
            Dim intIndex As Integer = 0

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor

            Dim strLink As String = ""
            Dim strImg As String = ""

            Try

                If strSortBy <> "" Then strSql = strSql & " '" & strSortBy & "','" & strDirection & "'"

                Trace.Write("DisplayReport ->strSql: " & strSql)

                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1



                    If isSortable Then

                        If strDirection = "ASC" Then

                            strDirection = "DESC"

                        Else 'If strDirection = "ASC" Then

                            strDirection = "ASC"

                        End If 'If strDirection = "ASC" Then


                        For intIndex = 0 To intColCount

                            qs.Clear()
                            qs("sort_by") = reader.GetName(intIndex)
                            qs("sort_dir") = strDirection

                            If strSortBy = reader.GetName(intIndex) Then

                                If strDirection = "ASC" Then

                                    strImg = "<img src=" & ProjectPath & "images/down.gif border=0>"

                                Else

                                    strImg = "<img src=" & ProjectPath & "images/up.gif border=0>"

                                End If

                            Else

                                strImg = ""

                            End If

                            strLink = "<a href=" & Request.Path & "?x=" & qs.ToString() & ">" & UCase(reader.GetName(intIndex)) & strImg & "</a>"

                            sb.Append("<td align=center valign=middle nowrap>" & strLink & "</td>")

                        Next

                    Else 'If isSortable Then

                        For intIndex = 0 To intColCount

                            sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                        Next

                    End If 'If isSortable Then

                    sb.Append("</tr>")

                    'REPORT ROWS
                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                        sb.Append("<tr class=GEN  " & strScript & ">")

                        For intIndex = 0 To reader.FieldCount() - 1

                            HttpContext.Current.Trace.Write(reader.GetName(intIndex) & ": " & reader(intIndex).GetType.ToString)

                            If Not IsDBNull(reader(intIndex)) Then

                                Select Case reader(intIndex).GetType.ToString

                                    Case "System.Double"

                                        sb.Append("<td valign=top align=right>" & FormatNumber(reader(intIndex), 2) & "</td>")

                                    Case "System.DateTime"

                                        sb.Append("<td valign=top align=center>" & Functions.FormatHispanicDateTime(reader(intIndex)) & "</td>")

                                    Case Else

                                        sb.Append("<td valign=top align=left>" & HttpContext.Current.Server.HtmlDecode(reader(intIndex)) & "</td>")

                                End Select

                            Else

                                sb.Append("<td valign=top align=left>N/A</td>")

                            End If

                        Next

                        sb.Append("</tr>")

                    Loop

                    sb.Append("</table>")


                Else 'If reader.HasRows() Then

                    sb.Append("<br>No encontramos la informaci&#243;n que solicitaste.")

                End If 'If reader.HasRows() Then

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function

        Public Function PageTemplate() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("<html>" & Chr(13))
                sb.Append("<head>" & Chr(13))

                sb.Append("<meta name=""author"" content=""Xavier Cabezas"">")

                sb.Append("<title>" & _objUser.getFirstName & " - Para Lideres-- Editor</title>" & Chr(13))

                'sb.Append("<script language=""JavaScript"" src=""" & _strProjectPath & "ajax.js"" type=""text/javascript""></script>" & vbLf)
                sb.Append("<script language=""JavaScript"" src=""" & _strProjectPath & "Ajax1.js"" type=""text/javascript""></script>" & vbLf)

                sb.Append("<link rel=""STYLESHEET"" href=""" & _strProjectPath & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""STYLESHEET"" href=""" & _strProjectPath & "Editor/editor.css"" type=""text/css"">" & vbLf)

                sb.Append("</head>" & Chr(13))

                sb.Append("<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & Chr(13))

                sb.Append("<div name=""divInstructions"" id=""divInstructions"" class='hideError' style=""FONT-FAMILY:Verdana;FONT-SIZE:9pt;BACKGROUND-COLOR:lightgoldenrodyellow;""></div>" & Chr(13))

                sb.Append("<div name='divPanel' id='divPanel' align=center valign=top class=panel>" & Chr(13))

                'sb.Append("<script language=""JavaScript"">" & Chr(13))

                'sb.Append("document.write(""<div align=center valign=top name='divAjaxContent' id='divAjaxContent' onmouseover='over=true;' onmouseout='over=false;' class=AjaxDivContent></div>"")" & Chr(13))

                sb.Append("<div align=center valign=top name='divAjaxContent' id='divAjaxContent' class=AjaxDivContent></div>" & Chr(13))

                'sb.Append("</script>" & Chr(13))

                sb.Append("</div>" & Chr(13))

                sb.Append("    <table width=""1000"" height=""600"" border=0 cellpadding=0 cellspacing=0 bgcolor=" & _strBgColor & ">")

                'TOP IMAGE
                'sb.Append("      <tr><td></td><td align=center valign=bottom align=center width=800 height=100><a href=editor.aspx><img src=" & _strProjectPath & "_images/header_contenido_large.jpg border=0></a></td><td></td></tr>" & Chr(13))
                sb.Append("      <tr bgcolor=#FEA803><td width=100>&nbsp;</td><td align=center valign=bottom width=800 height=100><center><a href=Editor/editor.aspx><img src=" & _strProjectPath & "images/header_contenido_large.jpg border=0></a></center></td><td width=100>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>" & Chr(13))


                sb.Append("      <tr><td colspan=3><br><br></td></tr>" & Chr(13))

                'SECOND ROW
                sb.Append("      <tr valign=""top"">" & Chr(13))

                'MENU WITH SECTIONS
                sb.Append("	        <td valign=top align=left with=200 class=SIDEMENU><br><br><br>" & Chr(13))
                sb.Append(EditorMenu())
                sb.Append("	        </td>" & Chr(13))


                'CENTER CELL
                sb.Append("	        <td valign=top align=center width=""880"" height=500>" & Chr(13))

                sb.Append(ContentLayout(_strPageTitle, _strPageContent))

                sb.Append("	        </td>" & Chr(13))

                'RIGHT CELL 
                sb.Append("	        <td valign=top align=left with=200 class=QUOTES>" & Chr(13))
                sb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                sb.Append("	        </td>" & Chr(13))

                sb.Append("      </tr>" & Chr(13))

                sb.Append("    </table> " & Chr(13))

                sb.Append("</body>" & Chr(13))
                sb.Append("</html>" & Chr(13))

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function ContentLayout(ByVal strTitle As String, ByVal strContent As String) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                sb.Append("<table width=""880"" border=0 cellpadding=0 cellspacing=0 bgcolor=white>")

                'FIRST ROW
                sb.Append("<tr>")
                sb.Append("<td bgcolor=" & _strBgColor & " valign=top width=60 height=60><img src=" & ProjectPath & "images/top-left.gif width=62 height=60></td>")
                sb.Append("<td class=TITLE valign=middle style=""border-top-color:black;border-top-width:2px;border-top-style:solid;"">" & strTitle & "&nbsp;</td>")
                sb.Append("<td bgcolor=" & _strBgColor & " valign=top width=60 height=60><img src=" & ProjectPath & "images/top-right.gif width=62 height=60></td>")
                sb.Append("</tr>" & Chr(13))

                'SECOND ROW
                sb.Append("<tr>")
                sb.Append("<td valign=top width=60 style=""border-left-color:black;border-left-width:2px;border-left-style:solid;border-bottom-color:black;border-bottom-width:2px;border-bottom-style:solid;"">&nbsp;</td>")

                sb.Append("<td  class=CONTENT height=600 valign=top style=""border-bottom-color:black;border-bottom-width:2px;border-bottom-style:solid;"">")
                sb.Append(strContent)
                sb.Append("</td>")

                sb.Append("<td valign=top width=60 style=""border-right-color:black;border-right-width:2px;border-right-style:solid;border-bottom-color:black;border-bottom-width:2px;border-bottom-style:solid;"">&nbsp;</td>")
                sb.Append("</tr>" & Chr(13))

                sb.Append("</table>")

                '--------------------

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function AddLink(ByVal strWebSection As String, ByVal intSectionId As Integer) As String

            Dim strLink As String = ""
            Dim strLink2 As String = ""

            Try

                Select Case strWebSection

                    Case "sections"

                        If intSectionId = 0 Then

                            'qs.Clear()
                            'qs("section_id") = intSectionId
                            'qs("web_section") = strWebSection
                            'qs("action") = "edit"

                            'strLink = "editor.aspx?x=" & qs.ToString()

                            'strLink = "<p><a href=" & ProjectPath & strLink & ">Añadir Nueva Secci&#243;n</a><p>"

                        Else

                            qs.Clear()
                            qs("section_id") = intSectionId
                            qs("web_section") = "articles"
                            qs("action") = "edit"

                            strLink = "<a href=" & ProjectPath & "Editor/editor.aspx?x=" & qs.ToString() & ">Añadir Nuevo Art&#237;culo Para Esta Secci&#243;n</a>"

                            qs.Clear()
                            qs("section_id") = 0
                            qs("parent_id") = intSectionId
                            qs("web_section") = "sections"
                            qs("action") = "edit"

                            strLink2 = "<a href=" & ProjectPath & "Editor/editor.aspx?x=" & qs.ToString() & ">Añadir Sub Secci&#243;n</a>"

                            strLink = "<p>" & strLink & " | " & strLink2 & "<p>"

                        End If

                    Case "users"

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "edit"

                        strLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strLink = "<p><a href=" & ProjectPath & strLink & ">Añadir Usuario</a><p>"

                    Case "cursos"

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "edit"

                        strLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strLink = "<p><a href=" & ProjectPath & strLink & ">Añadir Curso o Historieta</a><p>"

                    Case "surveys"

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "edit"

                        strLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strLink = "<p><a href=" & ProjectPath & strLink & ">Añadir Encuesta</a><p>"

                End Select

                Return strLink

            Catch ex As Exception

                Throw ex

            End Try

        End Function


        Public Function EditorContent(ByVal strWebSection As String, ByVal strAction As String, ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                Trace.Write("strAction: " & strAction)
                Trace.Write("strWebSection: " & strWebSection)
                Trace.Write("intSectionId: " & intSectionId)

                Select Case strAction

                    Case "main"

                        sb.Append(MainPage(strWebSection, intSectionId))

                    Case "edit"

                        Select Case strWebSection

                            Case "sections"

                                Dim intParentId As Integer = 0
                                Dim intThisSectionId As Integer = 0

                                Try
                                    intThisSectionId = CInt(qs("record_id"))
                                Catch
                                End Try

                                Try
                                    intParentId = CInt(qs("parent_id"))
                                Catch ex As Exception

                                End Try

                                sb.Append(GenEditSection(intThisSectionId, intParentId))

                            Case "articles"

                                Dim intPageId As Integer = 0

                                Try
                                    intPageId = CInt(qs("record_id"))
                                Catch
                                End Try

                                sb.Append(GenEditArticle(intPageId, intSectionId))

                            Case "users"

                                Dim intUserId As Integer = 0

                                Try
                                    intUserId = CInt(qs("record_id"))
                                Catch
                                End Try

                                sb.Append(GenEditUser(intUserId))

                            Case "cursos"

                                Dim intCursoId As Integer = 0

                                Try
                                    intCursoId = CInt(qs("record_id"))
                                Catch
                                End Try

                                sb.Append(GenEditCurso(intCursoId))

                            Case "surveys"

                                Dim intSurveyId As Integer = 0

                                Try
                                    intSurveyId = CInt(qs("record_id"))
                                Catch
                                End Try

                                sb.Append(GenEditSurvey(intSurveyId))

                            Case Else

                                sb.Append("En Construcci&#243;n")

                        End Select


                    Case "delete"

                        Select Case strWebSection

                            Case "sections"

                                Dim intThisSectionId As Integer = 0

                                Try
                                    intThisSectionId = CInt(qs("record_id"))
                                Catch
                                End Try

                                DeleteSection(intThisSectionId)

                            Case "articles"

                                Dim intPageId As Integer = 0

                                Try
                                    intPageId = CInt(qs("record_id"))
                                Catch
                                End Try

                                DeleteArticle(intPageId)

                            Case "users"

                                Dim intUserId As Integer = 0

                                Try
                                    intUserId = CInt(qs("record_id"))
                                Catch
                                End Try

                                DeleteUser(intUserId)

                            Case "cursos"

                                Dim intCursoId As Integer = 0

                                Try
                                    intCursoId = CInt(qs("record_id"))
                                Catch
                                End Try

                                DeleteCurso(intCursoId)

                            Case "surveys"

                                Dim intSurveyId As Integer = 0

                                Try
                                    intSurveyId = CInt(qs("record_id"))
                                Catch
                                End Try

                                DeleteSurvey(intSurveyId)

                            Case Else

                                sb.Append("En Construcci&#243;n")

                        End Select


                    Case "logoff"

                        Session.Abandon()

                        Response.Redirect("editor.aspx")

                    Case ""

                        sb.Append("Aqu&#237; va a poder manipular la informaci&#243;n de Para L&#237;deres.  Por favor elija una opci&#243;n del men&#250; de la izquierda")

                    Case Else

                        sb.Append("En Construcci&#243;n")

                End Select

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function MainPage(ByVal strWebSection As String, ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                Select Case strWebSection

                    Case "sections"

                        sb.Append(DisplayTableForSection(intSectionId))

                    Case "users"

                        sb.Append(SearchUser())

                    Case "cursos"

                        sb.Append(DisplayTableForCursos())

                    Case "surveys"

                        sb.Append(DisplayTableForSurveys())

                    Case Else

                        sb.Append("En Construcci&#243;n")

                End Select

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function GetSectionName(ByVal intSectionId As Integer) As String

            Return ParaLideres.GenericDataHandler.ExecScalar("SELECT section_name FROM sections WHERE section_id = " & intSectionId)

        End Function


        Public Function EditorMenu() As String

            Dim strMenuName As String = "EditorMenu" & _objUser.getSecurityLevel

            If IsNothing(Cache.Get(strMenuName)) Then

                Dim sb As New System.Text.StringBuilder("")

                Dim arrMenuLabels As String() = Split("cursos,encuestas,foros,secciones,usuarios", ",")
                Dim arrMenu As String() = Split("cursos,surveys,forum,sections,users", ",")
                Dim arrSec As String() = Split("6,6,6,4,6", ",")

                Dim intIndex As Integer = 0

                Dim strLink As String = ""

                Try

                    sb.Append("<ul>")

                    For intIndex = 0 To UBound(arrMenu)

                        If _objUser.getSecurityLevel >= CInt(arrSec(intIndex)) Then

                            qs.Clear()
                            qs("section_id") = CInt(Request("section_id"))
                            qs("web_section") = arrMenu(intIndex)
                            qs("action") = "main"

                            strLink = "Editor/editor.aspx?x=" & qs.ToString()

                            sb.Append("<li><a href=" & ProjectPath & strLink & ">" & UCase(arrMenuLabels(intIndex)) & "</a></li>")

                        End If

                    Next

                    qs.Clear()
                    qs("action") = "logoff"
                    sb.Append("<li><a href=" & ProjectPath & "Editor/editor.aspx?x=" & qs.ToString() & ">DESCONECTARSE</a></li>")

                    sb.Append("</ul>")

                    Cache.Insert(strMenuName, sb.ToString())

                Catch ex As Exception

                    Throw ex

                Finally

                    sb = Nothing

                End Try

            End If

            If Session("letmein") > 0 Then

                Return Cache.Get(strMenuName)

            Else

                Return "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

            End If

        End Function

        Public Function AlphaMenu(ByVal strIndex As String, ByVal strURL As String) As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")

            Dim x As Integer
            Dim strLink As String = ""

            sb.Append("<p align=center>")

            For x = 65 To 90
                If Asc(strIndex) <> x Then

                    qs.Clear()
                    qs("section_id") = 0
                    qs("web_section") = "addresses"
                    qs("action") = "main"
                    qs("alpha") = Chr(x)

                    strLink = strURL & "?x=" & qs.ToString()

                    sb.Append("<a href=" & strLink & " class=GEN>" & Chr(x) & "</a>&nbsp;")

                Else

                    sb.Append("<span class=BOLD>" & Chr(x) & "</b></span>&nbsp;")

                End If
            Next

            sb.Append("</p>")

            Return sb.ToString

            sb = Nothing

        End Function

        Public Function DisplayAddresses(ByVal strAlphaIndex As String, ByVal strWebSection As String, ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""

            Dim intRecordId As Integer = 0

            Dim strSQL As String = ""


            Try

                If strAlphaIndex = "" Then strAlphaIndex = "a"

                strSQL = "select * from addresses where last like '" & strAlphaIndex & "%' order by last, first"

                reader = ParaLideres.GenericDataHandler.GetRecords(strSQL)

                sb.Append(SetTableProperties(640, Rules.None))

                sb.Append("<tr><td colspan=2>" & AlphaMenu(strAlphaIndex, "Editor/editor.aspx") & "</td></tr>")

                If reader.HasRows() Then

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        sb.Append("<tr class=GEN  " & strScript & ">")

                        intRecordId = reader(0)

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "edit"
                        qs("record_id") = intRecordId

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<td valign=top align=center><a href=" & ProjectPath & strEditLink & ">EDIT</a></td>"

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "delete"
                        qs("record_id") = intRecordId

                        strDeleteLink = "delete.aspx?x=" & qs.ToString()

                        strDeleteLink = "<td valign=top align=center><a href=" & ProjectPath & strDeleteLink & ">DELETE</a></td>"

                        sb.Append("<tr class=BOLD " & strScript & "><td colspan=2>" & reader(1) & " " & reader(2) & " " & reader(4) & "</td></tr>")
                        sb.Append("<tr class=GEN " & strScript & "><td colspan=2>" & reader(5) & "</td></tr>")
                        sb.Append("<tr class=GEN " & strScript & "><td colspan=2>" & reader(6) & "</td></tr>")
                        sb.Append("<tr class=GEN " & strScript & "><td colspan=2>" & reader(7) & ", " & reader(8) & " " & reader(9) & "</td></tr>")
                        sb.Append("<tr class=GEN  " & strScript & "><td colspan=2>Phone: " & reader(10) & "</td></tr>")
                        sb.Append("<tr class=GEN  " & strScript & "><td colspan=2>E-mail: " & reader(11) & "</td></tr>")

                        sb.Append("<tr class=GEN  " & strScript & ">" & strEditLink & strDeleteLink & "</td></tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                Else

                    sb.Append("<tr><td colspan=2>No addresses were found.</td></tr>")

                End If

                sb.Append("</table>")

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function

        Public Function DisplayTableForSection(ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""
            Dim strOtherAction As String = ""
            Dim strSql As String = ""
            Dim strWebSection As String = "sections"
            Dim strSectionName As String = ""
            Dim strInMenu As String = ""

            Dim intIndex As Integer = 0
            Dim intRecordId As Integer = 0
            Dim intColCount As Integer = 0

            Dim blAllowChanges As Boolean = False

            Try

                'Allow everyone to add new articles for this section
                sb.Append(AddLink(strWebSection, intSectionId))

                If _objUser.getSecurityLevel > 4 Then

                    blAllowChanges = True

                    'strSql = "proc_GetSectionsByParentId " & intSectionId

                    strSql = "SELECT section_id, section_name AS 'Seccion', 'Aparece en el Menu' = CASE post_in_menu WHEN 0 THEN 'No' WHEN 1 THEN 'S&#237;' ELSE 'N/A' END FROM sections WHERE section_parent = " & intSectionId & " ORDER BY section_name"

                Else

                    If intSectionId > 0 Then

                        'TODO: change this
                        'strSql = "proc_GetSectionsByParentId " & intSectionId

                        strSql = "SELECT section_id, section_name AS 'Seccion', 'Aparece en el Menu' = CASE post_in_menu WHEN 0 THEN 'No' WHEN 1 THEN 'S&#237;' ELSE 'N/A' END FROM sections WHERE section_parent = " & intSectionId & " ORDER BY section_name"

                    Else

                        'TODO: make this an stored proc
                        strSql = "SELECT section_id, section_name AS 'Seccion', 'Aparece en el Menu' = CASE post_in_menu WHEN 0 THEN 'No' WHEN 1 THEN 'S&#237;' ELSE 'N/A' END FROM sections WHERE user_id = " & _objUser.getId & " ORDER BY section_name"

                    End If

                End If

                Trace.Write("strSql: " & strSql)

                'SHOW SUB SECTIONS
                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1

                    For intIndex = 1 To intColCount

                        sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                    Next

                    sb.Append("<td colspan=3 align=center valign=middle nowrap>ACCIONES</td>")

                    sb.Append("</tr>")

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intRecordId = reader(0)
                        strSectionName = reader(1)
                        strInMenu = reader(2)

                        If blAllowChanges Then

                            qs.Clear()
                            qs("section_id") = intSectionId
                            qs("web_section") = strWebSection
                            qs("action") = "edit"
                            qs("record_id") = intRecordId

                            strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                            strEditLink = "<td valign=top align=center><a href=" & ProjectPath & strEditLink & ">EDITAR</a></td>"

                            qs.Clear()
                            qs("section_id") = intSectionId
                            qs("web_section") = strWebSection
                            qs("action") = "delete"
                            qs("record_id") = intRecordId

                            strDeleteLink = "<td valign=top align=center><a href=""javascript:ConfirmDelete('Editor/editor.aspx?x=" & qs.ToString() & "' , 'esta secci&#243;n');"">BORRAR</a></td>"

                        End If

                        qs.Clear()
                        qs("section_id") = intRecordId 'this is to link to sections
                        qs("web_section") = "sections"
                        qs("action") = "main"
                        qs("record_id") = 0

                        strOtherAction = "Editor/editor.aspx?x=" & qs.ToString()

                        strOtherAction = "<td valign=top align=center><a href=" & ProjectPath & strOtherAction & ">VER</a></td>"

                        'CONTENT ROW
                        sb.Append("<tr class=GEN  " & strScript & ">")

                        sb.Append("<td valign=top align=left>" & HttpContext.Current.Server.HtmlDecode(strSectionName) & "</td>")

                        sb.Append("<td valign=top align=left>" & strInMenu & "</td>")

                        sb.Append(strEditLink)

                        sb.Append(strDeleteLink)

                        sb.Append(strOtherAction)

                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")

                    sb.Append(ConfirmDelete())

                End If 'if reader.hasrows()

                If intSectionId > 0 Then

                    sb.Append("<br><br>Art&#237;culos De Esta Secci&#243;n<br>")

                    sb.Append(DisplayTableForArticlesBySection(intSectionId))

                End If

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function

        Public Function DisplayTableForArticlesBySection(ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""
            Dim strOtherAction As String = ""
            Dim strSql As String = "proc_GetArticlesBySectionForEditor " & intSectionId
            Dim strWebSection As String = "articles"

            Dim strTitle As String = ""
            Dim strAuthor As String = ""
            Dim strAuthorId As String = ""
            Dim strAritlceType As String = ""

            Dim dtPosted As Date

            Dim intIndex As Integer = 0
            Dim intRecordId As Integer = 0
            Dim intColCount As Integer = 0

            Dim intIsPosted As Integer = 0
            Dim intIsFeatured As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1

                    For intIndex = 1 To intColCount

                        sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                    Next

                    sb.Append("<td colspan=3 align=center valign=middle nowrap>ACCIONES</td>")

                    sb.Append("</tr>")

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intRecordId = reader(0)
                        strTitle = reader(1)
                        dtPosted = reader(2)
                        intIsPosted = reader(3)
                        strAritlceType = reader(4)
                        intIsFeatured = reader(5)
                        strAuthor = reader(6)
                        strAuthorId = reader(7)

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "edit"
                        qs("record_id") = intRecordId

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<a href=" & ProjectPath & strEditLink & ">EDITAR</a>"

                        qs.Clear()
                        qs("section_id") = intSectionId
                        qs("web_section") = strWebSection
                        qs("action") = "delete"
                        qs("record_id") = intRecordId

                        strDeleteLink = "<a href=""javascript:ConfirmDelete('editor.aspx?x=" & qs.ToString() & "' , 'este art&#237;culo');"">BORRAR</a>"

                        'http://www.diandde.com/VerArticulo.aspx?Idp=1708&Ida=7717
                        'strOtherAction = "article.aspx?page_id=" & intRecordId
                        strOtherAction = "VerArticulo.aspx?Idp=" & intRecordId '& "&Ida=" & strAuthorId
                        'strOtherAction = "<a href=" & ProjectPath & strOtherAction & " target=new>VER</a>"
                        strOtherAction = "<a href=" & _strProjectPath1 & strOtherAction & " target=new>VER</a>"

                        'CONTENT ROW
                        sb.Append("<tr class=GEN  " & strScript & ">")
                        sb.Append(Functions.GenerateCell(HttpContext.Current.Server.HtmlDecode(strTitle)))
                        sb.Append(Functions.GenerateCell(dtPosted, Functions.ColType.DateType))
                        sb.Append(Functions.GenerateCell(intIsPosted, Functions.ColType.ImageCheck))
                        sb.Append(Functions.GenerateCell(strAritlceType))
                        sb.Append(Functions.GenerateCell(intIsFeatured, Functions.ColType.ImageCheck))
                        sb.Append(Functions.GenerateCell(strAuthor))
                        sb.Append(Functions.GenerateCell(strEditLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strDeleteLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strOtherAction, Functions.ColType.TextAlignCenter))
                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")

                    sb.Append(ConfirmDelete())

                Else 'If reader.HasRows() Then

                    sb.Append("<br>No encontramos art&#237;culos para esta secci&#243;n")

                End If 'If reader.HasRows() Then

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function
        Public Function DisplayTableForUsers(ByVal strParam As String, Optional ByVal strSortBy As String = "Nombre") As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intUserId As Integer = 0
            Dim intColCount As Integer = 0
            Dim intIndex As Integer = 0

            Dim strName As String = ""
            Dim strEmail As String = ""
            Dim strCountry As String = ""
            Dim strPassword As String = ""
            Dim strGender As String = ""
            Dim strUserType As String = ""

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""
            Dim strOtherAction As String = ""
            Dim strSql As String = "proc_SearchForUser '" & strParam & "', '" & strSortBy & "'"
            Dim strWebSection As String = "users"

            Dim dtLstVisit As Date

            Try

                sb.Append(AddLink(strWebSection, 0))

                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1

                    For intIndex = 1 To intColCount

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "main"
                        qs("last_name") = strParam
                        qs("sort_by") = reader.GetName(intIndex)

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<a href=" & ProjectPath & strEditLink & ">" & UCase(reader.GetName(intIndex)) & "</a>"

                        sb.Append("<td align=center valign=middle nowrap>" & strEditLink & "</td>")

                    Next

                    sb.Append("<td colspan=3 align=center valign=middle nowrap>ACCIONES</td>")

                    sb.Append("</tr>")

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intUserId = reader(0)
                        strName = reader(1)
                        strEmail = reader(2)
                        strPassword = reader(3)
                        strCountry = reader(4)
                        strGender = reader(5)
                        strUserType = reader(6)
                        dtLstVisit = reader(7)

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "edit"
                        qs("record_id") = intUserId

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<a href=" & ProjectPath & strEditLink & ">EDITAR</a>"

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "delete"
                        qs("record_id") = intUserId

                        strDeleteLink = "<a href=""javascript:ConfirmDelete('Editor/editor.aspx?x=" & qs.ToString() & "' , 'este usuario');"">BORRAR</a>"

                        strOtherAction = "<a href=" & ProjectPath & "Editor/emailpassword.aspx?Email=" & strEmail & "&send=yes target=new>ENVIAR CLAVE</a>"

                        'CONTENT ROW
                        sb.Append("<tr class=GEN  " & strScript & ">")

                        sb.Append(Functions.GenerateCell(strName))
                        sb.Append(Functions.GenerateCell(strEmail))
                        sb.Append(Functions.GenerateCell(strPassword))
                        sb.Append(Functions.GenerateCell(strCountry))
                        sb.Append(Functions.GenerateCell(strGender))
                        sb.Append(Functions.GenerateCell(strUserType))
                        sb.Append(Functions.GenerateCell(dtLstVisit, Functions.ColType.DateType))
                        sb.Append(Functions.GenerateCell(strEditLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strDeleteLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strOtherAction, Functions.ColType.TextAlignCenter))
                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")

                    qs.Clear()
                    qs("web_section") = strWebSection
                    qs("action") = "main"

                    strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                    strEditLink = "<a href=" & ProjectPath & strEditLink & ">Realizar Otra B&#250;squeda</a>"

                    sb.Append("<p align=center>" & strEditLink & "</p><br><br>")

                    sb.Append(ConfirmDelete())

                Else 'If reader.HasRows() Then

                    sb.Append("<br>No encontramos usuarios con el apellido " & strParam & " o con el e-mail: " & strParam)

                    qs.Clear()
                    qs("web_section") = strWebSection
                    qs("action") = "main"

                    strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                    strEditLink = "<a href=" & ProjectPath & strEditLink & ">Realizar Otra B&#250;squeda</a>"

                    sb.Append("<p align=center>" & strEditLink & "</p><br><br>")

                End If 'If reader.HasRows() Then

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function

        Public Function DisplayTableForCursos() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intCursoId As Integer = 0
            Dim intColCount As Integer = 0
            Dim intIndex As Integer = 0

            Dim strName As String = ""
            Dim strCursoDesc As String = ""
            Dim strCursoType As String = ""

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""
            Dim strOtherAction As String = ""
            Dim strSql As String = "proc_GetCursosForEditor"
            Dim strWebSection As String = "cursos"

            Try

                sb.Append(AddLink(strWebSection, 0))

                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1

                    For intIndex = 1 To intColCount

                        sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                    Next

                    sb.Append("<td colspan=3 align=center valign=middle nowrap>ACCIONES</td>")

                    sb.Append("</tr>")

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intCursoId = reader(0)
                        strName = reader(1)
                        If Not reader.IsDBNull(2) Then strCursoDesc = reader(2) Else strCursoDesc = ""
                        strCursoType = reader(3)

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "edit"
                        qs("record_id") = intCursoId

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<a href=" & ProjectPath & strEditLink & ">EDITAR</a>"

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "delete"
                        qs("record_id") = intCursoId

                        strDeleteLink = "<a href=""javascript:ConfirmDelete('Editor/editor.aspx?x=" & qs.ToString() & "' , 'este curso');"">BORRAR</a>"

                        'strOtherAction = "<a href=" & ProjectPath & "emailpassword.aspx?Email=" & strEmail & "&send=yes target=new>ENVIAR CLAVE</a>"

                        'CONTENT ROW
                        sb.Append("<tr class=GEN  " & strScript & ">")

                        sb.Append(Functions.GenerateCell(strName))
                        sb.Append(Functions.GenerateCell(strCursoDesc))
                        sb.Append(Functions.GenerateCell(strCursoType))
                        sb.Append(Functions.GenerateCell(strEditLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strDeleteLink, Functions.ColType.TextAlignCenter))
                        'sb.Append(Functions.GenerateCell(strOtherAction, Functions.ColType.TextAlignCenter))
                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")

                    sb.Append(ConfirmDelete())

                Else 'If reader.HasRows() Then

                    sb.Append("<br>No encontramos cursos el la base de datos")

                End If 'If reader.HasRows() Then

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function


        Public Function DisplayTableForSurveys() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intSurveyId As Integer = 0
            Dim intColCount As Integer = 0
            Dim intIndex As Integer = 0

            Dim strName As String = ""

            Dim dtFrom As Date
            Dim dtTo As Date

            Dim strScript As String = ""
            Dim strBgColor As String = _strBgColor
            Dim strEditLink As String = ""
            Dim strDeleteLink As String = ""
            Dim strOtherAction As String = ""
            Dim strSql As String = "proc_GetSurveysForEditor"
            Dim strWebSection As String = "surveys"

            Try

                sb.Append(AddLink(strWebSection, 0))

                reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                If reader.HasRows() Then

                    sb.Append(SetTableProperties())

                    'HEADER ROW
                    sb.Append("<tr class=BOLD bgcolor=" & _strBgColor & ">")

                    intColCount = reader.FieldCount() - 1

                    For intIndex = 1 To intColCount

                        sb.Append("<td align=center valign=middle nowrap>" & UCase(reader.GetName(intIndex)) & "</td>")

                    Next

                    sb.Append("<td colspan=3 align=center valign=middle nowrap>ACCIONES</td>")

                    sb.Append("</tr>")

                    Do While (reader.Read())

                        strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intSurveyId = reader(0)
                        strName = reader(1)
                        dtFrom = reader(2)
                        dtTo = reader(3)

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "edit"
                        qs("record_id") = intSurveyId

                        strEditLink = "Editor/editor.aspx?x=" & qs.ToString()

                        strEditLink = "<a href=" & ProjectPath & strEditLink & ">EDITAR</a>"

                        qs.Clear()
                        qs("web_section") = strWebSection
                        qs("action") = "delete"
                        qs("record_id") = intSurveyId

                        strDeleteLink = "<a href=""javascript:ConfirmDelete('Editor/editor.aspx?x=" & qs.ToString() & "' , 'esta encuesta');"">BORRAR</a>"

                        strOtherAction = "<a href=" & ProjectPath & "Editor/other_surveys.aspx?survey_id=" & intSurveyId & " target=new>VER</a>"

                        'CONTENT ROW
                        sb.Append("<tr class=GEN  " & strScript & ">")

                        sb.Append(Functions.GenerateCell(strName))
                        sb.Append(Functions.GenerateCell(dtFrom, ColType.DateType))
                        sb.Append(Functions.GenerateCell(dtTo, ColType.DateType))
                        sb.Append(Functions.GenerateCell(strEditLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strDeleteLink, Functions.ColType.TextAlignCenter))
                        sb.Append(Functions.GenerateCell(strOtherAction, Functions.ColType.TextAlignCenter))
                        sb.Append("</tr>")

                        If strBgColor = "white" Then strBgColor = _strBgColor Else strBgColor = "white"

                    Loop

                    sb.Append("</table>")

                    sb.Append(ConfirmDelete())

                Else 'If reader.HasRows() Then

                    sb.Append("<br>No encontramos encuestas el la base de datos")

                End If 'If reader.HasRows() Then

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function




        Private Function SearchUser() As String

            Dim strParam As String = ""
            Dim strSortBy As String = "Nombre"

            'Trace.Write("qs(last_name): " & qs("last_name"))
            'Trace.Write("qs(sort_by): " & qs("sort_by"))

            If Request.Form("last_name") <> "" Then

                strParam = Trim(Request.Form("last_name"))

            ElseIf qs("last_name") <> "" Then

                strParam = Trim(qs("last_name"))

            End If


            If Request("sort_by") <> "" Then

                strSortBy = Trim(Request("sort_by"))

            ElseIf qs("sort_by") <> "" Then

                strSortBy = Trim(qs("sort_by"))

            End If

            Trace.Write("strParam: " & strParam)
            Trace.Write("strSortBy: " & strSortBy)


            If strParam <> "" Then


                If strSortBy <> "" Then

                    Return DisplayTableForUsers(strParam, strSortBy)

                Else

                    Return DisplayTableForUsers(strParam)

                End If

            Else


                Dim sb As New System.Text.StringBuilder("")
                Dim objFrm As ParaLideres.FormControls.GenericForm

                Try

                    sb.Append(AddLink("users", 0))

                    objFrm = New ParaLideres.FormControls.GenericForm("FrmUsrSearch")

                    sb.Append(objFrm.FormAction("editor.aspx", 500))

                    sb.Append(objFrm.FormHidden("action", "main"))

                    sb.Append(objFrm.FormHidden("web_section", "users"))

                    sb.Append(objFrm.FormTextBox("Apellido o E-mail", "last_name", strParam, 40, "Ingresar el apellido paterno del usuario o e-mail para realizar la b&#250;squeda", True))

                    sb.Append(objFrm.FormEnd("Buscar", "last_name"))

                    sb.Append("<p align=center>|")

                    sb.Append("<a href=report_lastvisits.aspx>Ver Reporte de 30 D&#237;as</a>|")

                    sb.Append("<a href=show_bloqued.aspx>Usuarios Bloqueados</a>|")

                    sb.Append("</p><br><br>")

                    Return sb.ToString()

                Catch ex As Exception

                    Throw ex

                Finally

                    sb = Nothing
                    objFrm = Nothing

                End Try


            End If

        End Function

        Public Function ConfirmDelete(Optional ByVal strActionToTake As String = "borrar") As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                sb.Append(Chr(13) & "<script language=javascript>" & Chr(13))

                sb.Append("function ConfirmDelete(strQueryString, strDescription)" & Chr(13))

                sb.Append("{" & Chr(13))

                sb.Append("myvar = confirm('¡Advertencia! \n \n ¿Est&#225;s seguro de que quieres " & strActionToTake & " ' + strDescription + '?');" & Chr(13))

                sb.Append("if (myvar) {" & Chr(13))

                'sb.Append("alert(strQueryString); " & Chr(13))

                'sb.Append("myVarX = setTimeout('window.location.href=strQueryString',0); " & Chr(13))

                'sb.Append("setTimeout(Redirect(strQueryString),0);" & Chr(13))

                'sb.Append("window.location=strQueryString;" & Chr(13))

                sb.Append("window.setTimeout('window.location=""' + strQueryString + '""', 0);" & Chr(13))

                sb.Append("}" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("</script>" & Chr(13))

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Public Function SetTableProperties(Optional ByVal intWidth As Integer = 500, Optional ByVal intRules As Rules = 1) As String

            Dim style1 As String = "border-color:Black;border-width:1px;border-style:solid;width:"
            Dim style2 As String = "px;border-collapse:collapse"
            Dim strRules As String = "rows"

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")

            Try

                Select Case intRules

                    Case Rules.None

                        strRules = "none"

                        style1 = "border-color:Black;border-width:0px;border-style:solid;width:"
                        style2 = "px;border-collapse:collapse"

                    Case Rules.All

                        strRules = "all"

                    Case Rules.Cols

                        strRules = "columns"

                    Case Rules.Rows

                        strRules = "rows"

                End Select

                sb.Append("<table cellpadding=2 cellspacing=0 rules=""" & strRules & """ bordercolor=""Black"" border=""1"" style=""" & style1 & intWidth & style2 & """>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function GenEditArticle(ByVal intPageId As Integer, ByVal intSectionId As Integer) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm
            Dim sb As System.Text.StringBuilder
            Dim objPages As DataLayer.pages

            Dim intAuthor As Integer = 0

            Try

                sb = New System.Text.StringBuilder("")

                objFrm = New ParaLideres.FormControls.GenericForm("frmpages")

                objPages = New DataLayer.pages(intPageId)

                If intPageId > 0 Then

                    Me.PageTitle = "Editar Art&#237;culo: " & objPages.getPageTitle

                Else

                    Me.PageTitle = "Añadir Nuevo Art&#237;culo"

                End If

                If objPages.getSectionId > 0 Then intSectionId = objPages.getSectionId

                If objPages.getUserId > 0 Then intAuthor = objPages.getUserId Else intAuthor = CInt(Session("letmein"))

                sb.Append(objFrm.FormAction("post_pages.aspx", 640, True))

                sb.Append(objFrm.FormHidden("page_id", intPageId))

                sb.Append(objFrm.FormSelect("Publicar en Secci&#243;n", "section_id", intSectionId, "proc_GetAllSections", "Selecciona la secci&#243;n donde este art&#237;culo aparecer&#225;", True, "Seleccionar Secci&#243;n"))

                If objPages.getPosted = #1/1/1900# Then objPages.setPosted(Date.Today())

                sb.Append(objFrm.FormDateCal("Fecha de Publicaci&#243;n", "posted", objPages.getPosted, "Selecciona la fecha de cuando este art&#237;culo fue publicado.", True))

                sb.Append(objFrm.FormTextBox("T&#237;tulo", "page_title", objPages.getPageTitle, 200, "Entrar T&#237;tulo de Art&#237;culo", True))

                sb.Append(objFrm.FormTextArea("Resumen", "blurb", objPages.getBlurb, 5, 80, "El resumen del art&#237;culo aparecer&#225; en la p&#225;gina de la secci&#243;n a la cual este art&#237;culo pertenece.  M&#225;ximo 500 caracteres.", True, 500))

                sb.Append(objFrm.FormTextArea("Texto", "body", objPages.getBody, 20, 80, "El texto del art&#237;culo aparecer&#225; en la p&#225;gina del mismo como el texto principal siempre y cuando se seleccione Normal o P&#225;gina sin Menus en el Tipo de Art&#237;culo.", False, 20000))

                sb.Append(objFrm.FormSelect("Tipo de Art&#237;culo", "typeofarticle", objPages.getTypeofarticle, "sp_GetPageTypes", "Seleccionar tipo de art&#237;culo.  Normal: crea una p&#225;gina con los menues a los lados del art&#237;culo.  P&#225;gina Sin Menues: crea una p&#225;gina que no tiene los menues de los lados.  Estudio en Format Word: usado para cuando no en vez del texto del art&#237;culo existe un documento en formato Microsoft&reg; Word", True, "Seleccionar tipo de art&#237;culo"))

                sb.Append(objFrm.FormSelect("Autor", "user_id", intAuthor, "sp_GetUsersWhoPostedArticles", "Selecciona el nombre del autor de este art&#237;culo. Si existe la informaci&#243;n, aparecer&#225; el nombre y foto del autor de este art&#237;culo inmediatamente debajo del t&#237;tulo.", True, "Seleccionar Autor"))

                sb.Append(objFrm.FormTextBox("Palabras Clave", "keywords", objPages.getKeywords, 200, "Para ayudar la gente encontrar este art&#237;culo, escribe desde 1 hasta 10  palabras claves que est&#233;n relacionadas con este art&#237;culo. S&#243;lo usa palabras,  no frases. No uses comas entre palabras, sep&#225;ralas solo con espacios. (Ejemplo: Amor Misericordia Perd&#243;m) M&#225;ximo 100 Characteres", False))

                sb.Append(objFrm.FormSelect("Libro de la Biblia", "book", objPages.getBook, "sp_GetAllBooks", "Selecciona un libro de la Biblia si este art&#237;culo tiene que ver con uno libro espec&#237;fico. Si no esta relacionado directamente con un libro de la Biblia no tienes que seleccionar", False, "Seleccionar Libro"))

                sb.Append(objFrm.FormTextBox("Cap&#237;tulo", "chapter", objPages.getChapter, 3, "Si has seleccionado un libro de la Biblia enconces selecciona un cap&#237;tulo del libro que seleccionaste que tenga relaci&#243;n directa con este art&#237;culo.  Si este art&#237;culo no tiene que ver con un libro o cap&#237;tulo en especifico no tienes que seleccionar.  Ingresa el n&#250;mero 0 para indicar 'Todo el Libro'", False, 1, False))

                sb.Append(objFrm.FormSingleCheckbox("Publicar En Para L&#237;deres", "isposted", objPages.getIsposted, "Si hay un visto, este art&#237;culo aparecer&#225; en el sitio.  De otra manera los usuarios no podr&#225;n verlo.", False))

                sb.Append(objFrm.FormSingleCheckbox("Incluir En Destacados", "isfeatured", objPages.getIsfeatured, "Si hay un visto, este art&#237;culo aparecer&#225; dentro de los art&#237;culos destacados.", False))

                sb.Append(objFrm.FormSingleCheckbox("Mostrar Opci&#243;n de Evaluaci&#243;n", "rating", objPages.getRating, "Si hay un visto, los usuarios podran votar para evaluarlo", False))

                sb.Append(objFrm.FormSingleCheckbox("Notificar al Autor por E-mail", "notify", 0, "Si hay un visto, el sistema enviar&#225; un e-mail al autor de este art&#237;culo avis&#225;ndole que ha sido publicado en Para L&#237;deres", False))

                Select Case intSectionId

                    Case 3 'Noticias

                        sb.Append(objFrm.FormFile("Subir Foto", "File", 50, "Si quieres incluir una foto con esta noticia adjunta un archivo de tipo: JPG de 500 pixeles de ancho.", False, ".jpg"))

                    Case 335 'Blogs

                        'do nothing

                    Case Else

                        sb.Append(objFrm.FormFile("Subir Archivo", "File", 50, "Si tienes el documento en formato MS Word ® (.doc, .docx) o Adobe Acrobat® (.pdf) o Presentacion en Powerpoint(.ppt,.pptx) lo puedes colocar en nuestro sitio aqu&#237;.", False, ".doc,.docx,.pdf,.ppt,.pptx"))

                End Select

                sb.Append(objFrm.FormEnd())

                Return sb.ToString()

            Catch ex As Exception

                ErrorInfo(ex, "GenEditArticle")

                Throw ex

            Finally

                objFrm = Nothing
                sb = Nothing
                objPages = Nothing

            End Try

        End Function

        Private Function GenEditSection(ByVal intSectionId As Integer, ByVal intSectionParent As Integer) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm
            Dim sb As System.Text.StringBuilder
            Dim objSections As DataLayer.sections

            Dim arrYesNo As String() = Split("No,Si", ",")

            Try

                sb = New System.Text.StringBuilder("")

                objFrm = New ParaLideres.FormControls.GenericForm("frmsections")

                objSections = New DataLayer.sections(intSectionId)

                If intSectionId > 0 Then

                    Me.PageTitle = "Editar Secci&#243;n: " & objSections.getSectionName

                Else

                    Me.PageTitle = "Añadir Nueva Secci&#243;n"

                End If

                If intSectionParent > 0 Then objSections.setSectionParent(intSectionParent)

                sb.Append(objFrm.FormAction("post_section_editor.aspx", 640))

                sb.Append(objFrm.FormHidden("section_id", intSectionId))

                sb.Append(objFrm.FormTextBox("Nombre", "section_name", objSections.getSectionName, 75, "Ingresa el nombre de esta secci&#243;n", True))

                sb.Append(objFrm.FormSelect("Subsecci&#243;n De", "section_parent", objSections.getSectionParent, "proc_GetAllSections", "Seleccionar a que secci&#243;n pertenece", False, "Selecciona Secci&#243;n"))

                sb.Append(objFrm.FormTextArea("Descripci&#243;n Corta", "blurb", objSections.getBlurb, 5, 70, "Ingresa la descripci&#243;n corta acerca de esta secci&#243;n que aparece como resumen", True, 700))

                sb.Append(objFrm.FormTextArea("Descripci&#243;n Larga", "pagetext", objSections.getPagetext, 20, 70, "Ingresa la descripci&#243;n larga acerca de esta secci&#243;n que aparece como resumen", True, 4000))

                sb.Append(objFrm.FormSelect("Pertenece a", "user_id", objSections.getUserId, "SELECT id, upper(first_name + ' ' + last_name) FROM reg_users WHERE (security_level > 1 AND security_level < 7) ORDER BY first_name + ' ' + last_name", "Seleccionar el autor u organizaci&#243;n dueño de esta secci&#243;n", False, "Selecciona Autor"))

                sb.Append(objFrm.FormSelectArray("Mostrar en Menu", "post_in_menu", objSections.getPostInMenu, arrYesNo, "Seleccionar si quieres que esta secci&#243;n aparezca en el men&#250;", False, True))

                sb.Append(objFrm.FormEnd())

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objFrm = Nothing
                sb = Nothing
                objSections = Nothing

            End Try

        End Function

        Private Function GenEditCurso(ByVal intCursoId As Integer) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm
            Dim sb As System.Text.StringBuilder
            Dim objCursos As DataLayer.cursos

            Try

                sb = New System.Text.StringBuilder("")

                objFrm = New ParaLideres.FormControls.GenericForm("frmcursos")

                objCursos = New DataLayer.cursos(intCursoId)

                If intCursoId > 0 Then

                    Me.PageTitle = "Editar curso:" & objCursos.getTitle

                Else

                    Me.PageTitle = "Añadir curso"

                End If


                sb.Append(objFrm.FormAction("Editor/post_cursos.aspx", 640))

                sb.Append(objFrm.FormHidden("curso_id", intCursoId))

                sb.Append(objFrm.FormHidden("type", objCursos.getCursoType))

                sb.Append(objFrm.FormTextBox("Curso", "title", objCursos.getTitle, 100, "Inserte el nomber del curso", True))

                sb.Append(objFrm.FormTextArea("Descripci&#243;n", "curso_desc", objCursos.getCursoDesc, 5, 60, "Inserte la descripci&#243;n del curso", True, 1000))

                sb.Append(objFrm.FormTextBox("Versi&#243;n En L&#237;nea", "on_line", objCursos.getOnLine, 255, "Inserte donde est&#225; el curso para versi&#243;n en l&#237;nea", True))

                sb.Append(objFrm.FormTextBox("Versi&#243;n Para Descargar", "zip", objCursos.getZip, 255, "Inserte donde est&#225; el curso para versi&#243;n de descargar", True))

                'sb.Append(objFrm.FormTextBox("Type", "type", objCursos.getCursoType, 10, "Enter value for Type", True, 0, True))

                sb.Append(objFrm.FormEnd())

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objFrm = Nothing
                sb = Nothing
                objCursos = Nothing

            End Try


        End Function

        Private Function GenEditSurvey(ByVal intQuestionId As Integer) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm
            Dim sb As System.Text.StringBuilder
            Dim objQuestions As DataLayer.questions

            Dim dtStart As Date
            Dim dtEnd As Date

            Try

                sb = New System.Text.StringBuilder("")

                objFrm = New ParaLideres.FormControls.GenericForm("frmquestions")

                objQuestions = New DataLayer.questions(intQuestionId)

                If intQuestionId > 0 Then

                    Me.PageTitle = "Editar Encuesta: " & objQuestions.getQuestionDesc

                    dtStart = objQuestions.getDateStart
                    dtEnd = objQuestions.getDateEnd

                Else

                    Me.PageTitle = "Añadir Encuesta"

                    dtStart = CDate(ParaLideres.GenericDataHandler.ExecScalar("proc_GetLastDaySurveyPosted"))

                    dtStart = dtStart.AddDays(1)
                    dtEnd = dtStart.AddDays(30)

                End If

                sb.Append(objFrm.FormAction("post_survey_editor.aspx", 640))

                sb.Append(objFrm.FormHidden("question_id", intQuestionId))

                sb.Append(objFrm.FormTextBox("Encuesta", "question_desc", objQuestions.getQuestionDesc, 100, "Ingresar la pregunta para esta encuesta", True))

                sb.Append(objFrm.FormDateCal("Publicar Desde", "date_start", dtStart, "Seleccionar la fecha de inicio para publicar esta encuesta", True))

                sb.Append(objFrm.FormDateCal("Publicar Hasta", "date_end", dtEnd, "Seleccionar hasta cuando se debe publicar esta encuesta", True))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 1", "question_1", objQuestions.getQuestion1, 75, "Ingresar el texto para Opci&#243;n 1", True))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 2", "question_2", objQuestions.getQuestion2, 75, "Ingresar el texto para Opci&#243;n 2", True))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 3", "question_3", objQuestions.getQuestion3, 75, "Ingresar el texto para Opci&#243;n 3", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 4", "question_4", objQuestions.getQuestion4, 75, "Ingresar el texto para Opci&#243;n 4", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 5", "question_5", objQuestions.getQuestion5, 75, "Ingresar el texto para Opci&#243;n 5", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 6", "question_6", objQuestions.getQuestion6, 75, "Ingresar el texto para Opci&#243;n 6", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 7", "question_7", objQuestions.getQuestion7, 75, "Ingresar el texto para Opci&#243;n 7", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 8", "question_8", objQuestions.getQuestion8, 75, "Ingresar el texto para Opci&#243;n 8", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 9", "question_9", objQuestions.getQuestion9, 75, "Ingresar el texto para Opci&#243;n 9", False))

                sb.Append(objFrm.FormTextBox("Opci&#243;n 10", "question_10", objQuestions.getQuestion10, 75, "Ingresar el texto para Opci&#243;n 10", False))

                sb.Append(objFrm.FormEnd())

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objFrm = Nothing
                sb = Nothing
                objQuestions = Nothing

            End Try


        End Function

        Private Function GenEditUser(ByVal intUserId As Integer) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm
            Dim sb As System.Text.StringBuilder
            Dim objRegUsers As DataLayer.reg_users
            Dim flImage As System.IO.File

            Dim strPic As String = ""

            Try

                sb = New System.Text.StringBuilder("")

                objFrm = New ParaLideres.FormControls.GenericForm("frmreg_users")

                objRegUsers = New DataLayer.reg_users(intUserId)

                If intUserId > 0 Then

                    Me.PageTitle = "Editar Informaci&#243;n De " & objRegUsers.getFirstName & " " & objRegUsers.getLastName

                    strPic = ProjectPath & "files/" & objRegUsers.getPicture

                    If flImage.Exists(HttpContext.Current.Server.MapPath(strPic)) Then

                        sb.Append("<p align=center><img src=" & strPic & "></p>")

                    End If

                Else

                    Me.PageTitle = "Añadir Nuevo Usuario"

                End If

                '----------------------------------------------------------------------------------------------

                sb.Append(objFrm.FormAction("post_reg_users_editor.aspx", 640, True))

                sb.Append(objFrm.FormHidden("id", intUserId))

                'FirstName
                sb.Append(objFrm.FormTextBox("Nombre(s)", "first_name", objRegUsers.getFirstName, 50, "Inserte su nombre(s) (Luis, Luis Miguel, Maria, Maria Elena, Ma. Elena, etc.)", True))

                'LastName
                sb.Append(objFrm.FormTextBox("Apellido(s)", "last_name", objRegUsers.getLastName, 50, "Inserte su apellido(s) (Le&#243;n, Le&#243;n Gonzalez, etc.)", True))

                If intUserId > 0 Then

                    'Password
                    sb.Append(objFrm.FormPasswordTextBox("Clave Secreta", "Password", objRegUsers.getPassword, 16, "Inserte su clave secreta. M&#237;nimo 4 letras y m&#225;ximo 16 letras.", True, 4, False))

                Else

                    sb.Append(objFrm.FormHidden("Password", "Usr" & Date.Today.Second & "" & Date.Today.Day))

                End If

                'Email
                sb.Append(objFrm.FormTextBox("Email", "Email", objRegUsers.getEmail, 50, "Inserte su direcci&#243;n de E-mail", True, 5, False, True))

                'Bday
                sb.Append(objFrm.FormDateCal("Fecha de Nacimiento", "Bday", objRegUsers.getBday.Date, "Seleccionar D&#237;a de Nacimiento", True))

                'Sex
                sb.Append(objFrm.FormSelect("Sexo", "Sex", objRegUsers.getSex, "sp_GetAllSexo", "Selecciona Masculino o Femenino", True))

                'MStatus
                sb.Append(objFrm.FormSelect("Estado Civil", "m_status", objRegUsers.getMStatus, "sp_GetAllEstadoCivil", "Selecciona tu estado civil", True))


                'WorkType
                sb.Append(objFrm.FormSelect("Tipo de Trabajo", "work_type", objRegUsers.getWorkType, "sp_GetAllWorkTypes", "Selecciona tu tipo de trabajo con j&#243;venes", True))

                'City
                sb.Append(objFrm.FormTextBox("Ciudad", "City", objRegUsers.getCity, 50, "La ciudad donde vives", True))

                'State
                sb.Append(objFrm.FormTextBox("Estado", "State", objRegUsers.getState, 100, "El estado/provincia donde vives", False))

                'Country
                sb.Append(objFrm.FormSelect("Pa&#237;s", "Country", objRegUsers.getCountry, "sp_GetAllCountries", "Selecciona el pa&#237;s donde vives", True))

                'MainLanguage
                Dim arrLenguas As String()
                arrLenguas = Split("Español|Portugu&#233;s|Ingl&#233;s|Otro", "|")

                sb.Append(objFrm.FormSelectArray("Lenguaje Principal", "main_language", objRegUsers.getMainLanguage, arrLenguas, "Selecciona tu lenguage principal", True, False))

                'Phone
                sb.Append(objFrm.FormTextBox("Tel&#233;fono", "Phone", objRegUsers.getPhone, 15, "Tu n&#250;mero de tel&#233;fono", False))

                'SecurityLevel
                sb.Append(objFrm.FormSelect("Tipo de Usuario", "security_level", objRegUsers.getSecurityLevel, "sp_GetSecurityLevel", "Seleccionar el tipo de usuario.", True))

                'Picture
                sb.Append(objFrm.FormFile("Foto", "Picture", 50, "Sube la foto de este usuario en formato JPG. Foto debe ser de m&#225;ximo 200 pixeles de ancho.", False))

                'Otherinfo
                sb.Append(objFrm.FormTextArea("Biograf&#237;a", "otherinfo", objRegUsers.getOtherinfo, 20, 70, "Ingresa informaci&#243;n acerca de t&#237; que quieras compartir con otros miembros de Para L&#237;deres.", False, 400))

                sb.Append(objFrm.FormEnd())

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objFrm = Nothing
                sb = Nothing
                objRegUsers = Nothing

            End Try

        End Function

        Private Sub DeleteSurvey(ByVal intQuestionId As Integer)

            Dim objQuestions As DataLayer.questions

            Try

                objQuestions = New DataLayer.questions(intQuestionId)

                If intQuestionId > 0 Then

                    objQuestions.Delete(intQuestionId)

                End If

                qs("web_section") = "surveys"
                qs("action") = "main"

                Response.Redirect("Editor/editor.aspx?x=" & qs.ToString())

            Catch ex As Exception

                ShowError(ex)

            Finally

                objQuestions = Nothing

            End Try


        End Sub

        Private Sub DeleteCurso(ByVal intCursoId As Integer)

            Dim objCursos As DataLayer.cursos

            Try

                objCursos = New DataLayer.cursos(intCursoId)

                If intCursoId > 0 Then

                    objCursos.Delete(intCursoId)

                End If

                qs("web_section") = "cursos"
                qs("action") = "main"

                Response.Redirect("Editor/editor.aspx?x=" & qs.ToString())

            Catch ex As Exception

                ShowError(ex)

            Finally

                objCursos = Nothing

            End Try

        End Sub

        Private Sub DeleteSection(ByVal intSectionId As Integer)

            Dim objSection As DataLayer.sections

            Try

                objSection = New DataLayer.sections(intSectionId)

                qs("section_id") = objSection.getSectionParent
                qs("web_section") = "sections"
                qs("action") = "main"

                If intSectionId > 0 Then

                    objSection.Delete(intSectionId)

                End If

                Response.Redirect("Editor/editor.aspx?x=" & qs.ToString())

            Catch ex As Exception

                ShowError(ex)

            Finally

                objSection = Nothing

            End Try

        End Sub

        Private Sub DeleteArticle(ByVal intPageId As Integer)

            Dim objPages As DataLayer.pages

            Try

                objPages = New DataLayer.pages(intPageId)

                qs("section_id") = objPages.getSectionId
                qs("web_section") = "sections"
                qs("action") = "main"

                If intPageId > 0 Then

                    Select Case objPages.getSectionId

                        Case 3 'Noticias

                            Try
                                Cache.Remove("RSSFeedNoticias")
                            Catch ex As Exception
                            End Try

                        Case 335 'Blogs

                            Try
                                Cache.Remove("RSSFeedBlogs")
                            Catch ex As Exception
                            End Try

                    End Select

                    objPages.Delete(intPageId)

                End If

                Response.Redirect("editor.aspx?x=" & qs.ToString())

            Catch ex As Exception

                ShowError(ex)

            Finally

                objPages = Nothing

            End Try

        End Sub


        Private Sub DeleteUser(ByVal intUserId As Integer)

            Dim objRegUsers As DataLayer.reg_users

            Try

                objRegUsers = New DataLayer.reg_users(intUserId)

                qs("section_id") = 0
                qs("web_section") = "users"
                qs("action") = "main"

                If InStr(objRegUsers.getLastName, " ") > 0 Then

                    qs("last_name") = Left(objRegUsers.getLastName, InStr(objRegUsers.getLastName, " "))

                Else

                    qs("last_name") = objRegUsers.getLastName

                End If

                If intUserId > 0 Then

                    objRegUsers.Delete(intUserId)

                End If

                Response.Redirect("Editor/editor.aspx?x=" & qs.ToString())

            Catch ex As Exception

                ShowError(ex)

            Finally

                objRegUsers = Nothing

            End Try



        End Sub



        Public Function LogonForm() As String

            Dim sb As StringBuilder = New StringBuilder("")

            sb.Append("<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""6"">")
            sb.Append("	<tr>")
            sb.Append("		<td class=""blackheader"" width=""100%"">Por favor ingrese su informaci&#243;n</td>")
            sb.Append("	</tr>")
            sb.Append("	<tr>")
            sb.Append("		<td class=""maintext"" width=""100%""></td>")
            sb.Append("	</tr>")
            sb.Append("	<tr>")
            sb.Append("		<td class=""lgtheadersm"" width=""100%"">Si tiene problemas contacte a <a href=mailto:apoyo@ParaLideres.org>apoyo@ParaLideres.org</a></td>")
            sb.Append("	</tr>")
            sb.Append("</table>")
            sb.Append("")
            sb.Append("					")
            sb.Append("")
            sb.Append("<br>")
            sb.Append("<table border=""0"" bgcolor=""" & _strBgColor & """>")
            sb.Append("<tr>")
            sb.Append("	<td width=""100%"" bgcolor=""white"">")
            sb.Append("		<table border=""0"" width=""350"" cellspacing=""1"" cellpadding=""5"">")
            sb.Append("			<tr>")
            'sb.Append("				<td class=""blackheader"" width=""100%"" bgcolor=""" & _strBgColor & """ align=""right""><img src=""images/logon.gif"" alt=""LOGON""></td>")

            sb.Append("				<td class=""blackheader"" width=""100%"" bgcolor=""" & _strBgColor & """ align=""right"">INGRESAR</td>")

            sb.Append("			</tr>")
            'sb.Append("			<tr>")
            'sb.Append("				<td class=""maintext"" width=""100%"">Please enter your e-mail address and password information and then click on the &quot;LOGON&quot; Button.&nbsp; If you have forgotten your password, then click on the link below the logon button.</td>")
            'sb.Append("			</tr>")
            sb.Append("			<tr>")
            sb.Append("				<td class=""maintext"" width=""100%"">")
            sb.Append("					<form id=""form1"" name=""form1"" action=""checklogon.aspx"" method=""post"">")
            sb.Append("					<table border=""0"" bgcolor=""#9fabb3"" cellspacing=""0"" cellpadding=""0"" width=""100%"">")
            sb.Append("						<tr>")
            sb.Append("							<td width=""100%"" bgcolor=""white"">")
            sb.Append("								<table border=""0"" width=""100%"" cellpadding=""0"">")
            sb.Append("									<tr>")
            sb.Append("										<td class=""lgtheadersm"" width=""100%"" bgcolor=""" & _strBgColor & """ align=""right"" valign=""middle""><strong>USUARIO</strong>:&nbsp;&nbsp;</td>")
            sb.Append("										<td valign=""middle"" align=""left""><input maxlength=""100"" size=""30"" name=""username"" value="""" style=""font-family: Verdana, arial; font-size: 9pt; font-weight: normal; color: rgb(128,128,128); vertical-align: middle; border: 1px solid rgb(217,222,225); margin: 4px; padding: 2px"" type=""text""></td>")
            sb.Append("									</tr>")
            sb.Append("									<tr>")
            sb.Append("										<td class=""lgtheadersm"" width=""100%"" bgcolor=""" & _strBgColor & """ align=""right"" valign=""middle""><strong>CLAVE</strong>: &nbsp;</td>")
            sb.Append("										<td valign=""middle"" align=""left""><input type=""password"" maxlength=""16"" size=""30"" name=""password"" style=""font-family: Verdana, arial; font-size: 9pt; font-weight: normal; color: rgb(128,128,128); vertical-align: middle; border: 1px solid rgb(217,222,225); margin: 4px; padding: 2px""></td>")
            sb.Append("									</tr>")
            sb.Append("									<tr>")
            sb.Append("										<td colspan=""2"" align=""center"">")
            sb.Append("											<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""5"">")
            sb.Append("												<tr>")
            sb.Append("													<td width=""50%""></td>")
            sb.Append("													<td width=""50%"" valign=""middle"" align=""center""><input type=""submit"" value=""Entrar"" name=""submit"" style=""font-family: Verdana, arial; font-size: 9pt; background-color: rgb(217,222,225); border: 1px solid rgb(159,171,179); padding: 3px"" class=""lgtheadersm""></td>")
            sb.Append("												</tr>")
            sb.Append("											</table>")
            sb.Append("										</td>")
            sb.Append("									</tr>")
            sb.Append("								</table>")
            sb.Append("							</td>")
            sb.Append("						</tr>")
            sb.Append("					</table>")
            sb.Append("					</form>")
            sb.Append("				</td>")
            sb.Append("				</tr>")
            'sb.Append("				<tr>")
            'sb.Append("					<td class=""maintext"" width=""100%"" align=""center""><a href=""/email_psswrd.aspx"" class=""BOLD"">Did you forget your password?</a></td>")
            'sb.Append("				</tr>")
            sb.Append("			</table>")
            sb.Append("		</td>")
            sb.Append("	</tr>")
            sb.Append("</table>")

            'sb.Append(DataHandler.ConnectionString)

            Return sb.ToString()

            sb = Nothing

        End Function

        Public Function ErrorLayout(ByVal strErrorMessage As String) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                sb.Append("<table border=""0"" cellspacing=""0"" cellpadding=""0"">")
                sb.Append("<tr>")
                sb.Append("<td valign=""top"">")
                sb.Append("<img src=""images/error1.jpg"" alt=""Error""></td>")
                sb.Append("<td class=""duedate"" valign=""top"" width=""205"">")
                sb.Append("<img src=""images/error2.gif"" alt=""Error"">")
                sb.Append("<p>" & strErrorMessage & "</p>")
                sb.Append("</td>")
                sb.Append("</tr>")
                sb.Append("<tr>")
                sb.Append("<td bgcolor=""#BEBEBE"" height=""1""></td>")
                sb.Append("<td bgcolor=""#BEBEBE"" height=""1""></td>")
                sb.Append("</tr>")
                sb.Append("</table>")

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


#End Region

#Region "Subs"

        Sub New()


        End Sub

        Public Sub ShowError(ByVal excError As Exception, Optional ByVal strProcedureName As String = "N/A")

            Select Case excError.Message

                Case "Thread was being aborted.", "Could not access 'CDO.Message' object."

                    'do nothing ignore error

                Case Else

                    Functions.ShowError(excError, strProcedureName)

            End Select

        End Sub

        Public Sub ErrorInfo(ByVal ex As Exception, ByVal strProcName As String)

            Trace.Warn("Error on " & strProcName)
            Trace.Warn(ex.Source)
            Trace.Warn(ex.Message)
            Trace.Warn(ex.ToString)

        End Sub

#Region "Page Cycle"


        Protected Overrides Sub Construct()

            Trace.Write(_blIsDebugMode)

            Try

                EnableViewState = False
                Trace.IsEnabled = True

                Trace.Write("CONSTRUCT")

                Trace.Write(_blIsDebugMode)

            Catch exc As Exception

                ShowError(exc)

            End Try

        End Sub

        Protected Overrides Sub OnInit(ByVal e As System.EventArgs)

            Trace.Write("INIT")

            Dim intUserId As Integer = 0

            Dim intFoundBloqued As Integer = 0

            Dim script As String = System.IO.Path.GetFileName(Me.Request.FilePath)


            Try
                intFoundBloqued = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_GetBloquedIPAddress '" & Session("IPAddress") & "'"))
            Catch ex As Exception
            End Try


            If intFoundBloqued > 0 Then Response.Redirect("http://" & Session("IPAddress"))


            Trace.Write("script: " & script)

            If _blIsDebugMode Then Session("letmein") = 1 '12 profe jaime

            Try
                intUserId = CInt(Session("letmein"))
            Catch ex As Exception
            End Try

            _objUser = New DataLayer.reg_users(intUserId)

            Trace.IsEnabled = _blIsDebugMode

            If _objUser.getId = 1 Then Trace.IsEnabled = True

            Me.Server.ScriptTimeout = 600

            If Request("x") <> "" Then

                Try

                    qs = New SecureQueryString(Request("x"))

                Catch ex As Exception

                    qs = New SecureQueryString

                End Try

            Else

                qs = New SecureQueryString

            End If


            If IsNothing(Cache.Get("BGColorEditor")) Then

                Cache.Insert("BGColorEditor", _strBgColor)

            End If

            _strBgColor = Cache.Get("BGColorEditor")

            If intUserId < 1 And script <> "editor.aspx" Then Response.Redirect("Editor/editor.aspx")

        End Sub

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

            Trace.Write("RENDER")

            If _objUser.getId < 1 Then

                Me.PageTitle = "INGRESAR AL EDITOR DE PARA LIDERES"
                Me.PageContent = LogonForm()

            End If

            writer.Write(PageTemplate())

        End Sub

        Protected Overrides Sub OnUnload(ByVal e As System.EventArgs)

            qs = Nothing

            _objUser = Nothing

        End Sub

#End Region

#End Region

#Region "Security"


#End Region


    End Class


End Namespace
