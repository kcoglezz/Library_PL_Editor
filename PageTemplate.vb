Imports System
Imports System.Web
Imports System.Web.Mail
Imports System.Text
Imports System.Web.Caching
Imports System.Web.UI

Namespace ParaLideres

    Public Class PageTemplate

        Inherits Page

#Region "Declarations"

        Public _objUser As DataLayer.reg_users

        Public _arrLastViewed(9, 3) As String

        Public intDataGridWidth As Integer = 500
        Public _support_account As String = System.Web.Configuration.WebConfigurationManager.AppSettings("SupportAccount")
        Public _develop As String = System.Web.Configuration.WebConfigurationManager.AppSettings("DeveloperAccount")

        Private _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _downloads_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("DownloadsPath")
        Private _is_under_construction As Boolean = System.Web.Configuration.WebConfigurationManager.AppSettings("IsUnderConstruction")

        Private _page_content As String
        Private _page_title As String
        Private _section As String
        Private _query_collection As String
        Private _page_path As String = ""
        Private _strOnLoadScript As String = ""
        Private _strScript As String = ""
        Private _strBorderColor As String = "#D09000" 'orange
        Private _strBackGroundColor As String = "#A7D168" 'green

        Private _intWidth As Integer = 160 'width of left column in pixels
        Private _user_id As Integer = 0
        Private _security_level As Integer = 1
        Private intDataGridPagingSize As Integer = 30
        Private _intPageFormat As PageFormats = PageFormats.NormalPage

        Private _simple_page As Boolean = False
        Private _blRequiresLogin As Boolean = False


        Public Enum PageFormats As Integer

            NormalPage = 0
            ExcelFormat = 1
            WordFormat = 2
            PrintFormat = 3
            Email = 4

        End Enum

#End Region

#Region "Properties"

        Public Property ProjectPath() As String
            Get

                Return _project_path

            End Get

            Set(ByVal Value As String)

                _project_path = Value

            End Set

        End Property

        Public WriteOnly Property IsSimplePage() As Boolean

            Set(ByVal Value As Boolean)

                _simple_page = Value

            End Set

        End Property


        Public Property PageContent() As String

            Get

                Return _page_content

            End Get

            Set(ByVal Value As String)

                _page_content = Value

            End Set

        End Property

        Public Property PageTitle() As String

            Get

                Return _page_title

            End Get

            Set(ByVal Value As String)

                _page_title = Value

            End Set

        End Property

        Public WriteOnly Property Section() As String

            Set(ByVal Value As String)

                _section = Value

            End Set

        End Property

        Public Property PageSize() As Integer

            Get

                Return intDataGridPagingSize

            End Get

            Set(ByVal Value As Integer)

                intDataGridPagingSize = Value

            End Set

        End Property



        Public Property SecurityLevel() As Integer

            Get

                Return _security_level

            End Get

            Set(ByVal Value As Integer)

                _security_level = Value

            End Set

        End Property


        Public WriteOnly Property PagePath() As String
            Set(ByVal Value As String)

                _page_path = Value

            End Set
        End Property

        Public Property OnLoadScript() As String
            Get

                Return _strOnLoadScript

            End Get
            Set(ByVal Value As String)

                _strOnLoadScript = Value

            End Set
        End Property

        Public WriteOnly Property RequiresLogin() As Boolean

            Set(ByVal value As Boolean)

                _blRequiresLogin = value

            End Set

        End Property

#End Region

#Region "Functions"

        Public Function DisplayNews() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intPageId As Integer = 0

            Dim strImagePath As String = ""
            Dim strImageName As String = ""
            Dim strSummary As String = ""
            Dim strAlign As String = "left"
            Dim strTitle As String = ""
            Dim dtPosted As Date = #1/1/0900#

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetNewsForHomePage")

                If reader.HasRows Then

                    Do While reader.Read()

                        intPageId = reader(0)
                        strTitle = reader(1)
                        strSummary = reader(2)
                        dtPosted = reader(3)

                        strImageName = "pic_" & intPageId & "_sml.jpg"

                        strImagePath = Server.MapPath(ProjectPath & "files\") & strImageName

                        sb.Append("<p>")

                        sb.Append(strTitle & "<br>")

                        sb.Append("Publicado el " & Functions.FormatHispanicDateTime(dtPosted) & "<br>")

                        If System.IO.File.Exists(strImagePath) Then

                            sb.Append("<img src=" & ProjectPath & "files/" & strImageName & " align=""" & strAlign & """>")

                        End If

                        sb.Append(ParaLideres.Functions.ReplaceCR(strSummary))

                        sb.Append("</p>")

                        If strAlign = "left" Then

                            strAlign = "right"

                        Else

                            strAlign = "left"

                        End If

                    Loop

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


        Public Function DisplayBlogs() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intPageId As Integer = 0
            Dim intUserId As Integer = 0

            Dim strImageName As String = ""
            Dim strTitle As String = ""
            Dim strAuthor As String = ""
            Dim strSummary As String = ""

            Dim dtPosted As Date = #1/1/0900#

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetBlogsForHomePage")

                If reader.HasRows Then

                    Do While reader.Read()

                        intPageId = reader(0)
                        strTitle = reader(1)
                        strSummary = reader(2)
                        dtPosted = reader(3)
                        intUserId = reader(4)
                        strAuthor = reader(5)
                        strImageName = reader(6)

                        sb.Append("<p>")

                        sb.Append(Functions.ShowPicture(intUserId, strImageName))

                        sb.Append(strTitle & "<br>")

                        sb.Append("Publicado el " & Functions.FormatHispanicDateTime(dtPosted) & " " & "Por:" & strAuthor & "<br>")

                        sb.Append(ParaLideres.Functions.ReplaceCR(strSummary))

                        sb.Append("</p>")

                    Loop

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

        Public Function GetAllSubsections(ByVal intSectionId As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intValue As Integer

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("Select employee_id from employee where supervisor = " & intSectionId)

                If reader.HasRows Then

                    Do While reader.Read()

                        If Not reader.IsDBNull(0) Then intValue = reader(0)

                        sb.Append(intValue & ",")

                    Loop

                    sb.Remove(sb.Length - 1, 1)

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

        Private Function ShowLoading() As String

            Dim sbContent As StringBuilder = New StringBuilder("")

            sbContent.Append("<html><head><title>" & _page_title & "</title>" & vbLf)
            sbContent.Append("<script language=""JavaScript1.2"" src=""" & _project_path & "jscr.js"" type=""text/javascript""></script>" & vbLf)

            sbContent.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
            sbContent.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)

            'sbContent.Append("<script>" & Chr(13))
            'sbContent.Append("<!--" & Chr(13))
            'sbContent.Append("//function to display loading... message" & Chr(13))
            'sbContent.Append("function HideLoadingMessage(){" & Chr(13))
            'sbContent.Append("document.getElementById('LoadingDiv').style.visibility = 'hidden';" & Chr(13))
            'sbContent.Append("}" & Chr(13))
            'sbContent.Append("//-->" & Chr(13))
            'sbContent.Append("</script>" & Chr(13))

            sbContent.Append("</head>" & Chr(13))

            sbContent.Append("<body topmargin=""0"" leftmargin=""0"" ")

            'sbContent.Append(" onLoad=""HideLoadingMessage();""")

            sbContent.Append(" onLoad=""" & _strOnLoadScript & """")

            sbContent.Append(">" & Chr(13))

            '--------------------------------------------------------------------------
            'show this while loading page
            'sbContent.Append("<div id=LoadingDiv name=LoadingDiv align=center valign=center style=""position: absolute; left:0px; top:0px;background-color:#FFFFFF; layer-background-color:#FFFFFF;width:100%; height:100%; border-color:black;border-style:solid;border-width:1;"">")
            'sbContent.Append("<br><br><br><br><br><br><img src=" & _project_path & "images/progress.gif><br><b>Loading...</b>")
            'sbContent.Append("</div>")
            '--------------------------------------------------------------------------

            Return sbContent.ToString

        End Function


        Private Function Search() As String

            Dim sb As New System.Text.StringBuilder("")

            Dim strFormName As String = "frmSearch"
            Dim strField As String = "shearchparam"


            Try

                sb.Append("      <tr>")
                sb.Append("        <td><form name=""frmSearch"" method=""post"" action=""search.aspx"">" & vbCrLf)

                'sb.Append("<td><form name=""frmSearch"" action=""http://search.freefind.com/find.html"" method=""get"" accept-charset=""utf-8"" target=""_self"">" & vbCrLf)

                sb.Append("          <table width=""515"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
                sb.Append("            <tr>")
                sb.Append("              <td width=""156"" align=""center""><img src=""" & _project_path & "_images/buscar_por_palabra.gif"" width=""105"" height=""15""></td>")
                sb.Append("              <td width=""145""><input type=""text"" name=""" & strField & """ id=""" & strField & """ CLASS=frmInput>")

                'Search Paramaters
                'sb.Append("		<input type=""hidden"" name=""si"" value=""68396942"">" & vbCrLf)
                'sb.Append("		<input type=""hidden"" name=""pid"" value=""r"">" & vbCrLf)
                'sb.Append("		<input type=""hidden"" name=""n"" value=""0"">" & vbCrLf)
                'sb.Append("		<input type=""hidden"" name=""_charset_"" value="""">" & vbCrLf)
                'sb.Append("		<input type=""hidden"" name=""bcd"" value=""&#247;"">" & vbCrLf)
                'sb.Append("		<input type=""hidden"" name=""lang"" value=""es"">" & vbCrLf)

                sb.Append("</td>")

                sb.Append("              <td width=""145"">&nbsp;&nbsp;<input type=""button""  value=""Buscar"" onclick='VerifyForm" & strFormName & "();' name=""btn" & strFormName & """ id=""btn" & strFormName & """ class=frmbutton></td>")

                sb.Append("            </tr>")
                sb.Append("          </table>")
                sb.Append("        </form></td>")
                sb.Append("      </tr>")

                sb.Append("<script language=javascript>" & Chr(13))
                sb.Append("<!--" & Chr(13))

                'function VerifyForm()
                sb.Append(Chr(13) & "function VerifyForm" & strFormName & "(){" & Chr(13))

                sb.Append("var buttonclicks = 0;" & Chr(13))

                sb.Append("if (document." & strFormName & "." & strField & ".value == """"){" & Chr(13))
                sb.Append("	alert('Por favor ingresa una palabra para tu búsqueda');" & Chr(13))
                sb.Append("document." & strFormName & "." & strField & ".focus();" & Chr(13))
                sb.Append("document." & strFormName & "." & strField & ".select();" & Chr(13))
                sb.Append("}" & Chr(13))

                'Request at least 4 characters for search
                sb.Append("else if (GetFieldLength(document." & strFormName & "." & strField & ") < 4){" & Chr(13))
                sb.Append("	alert('Por favor ingresa una palabra de al menos cuatro letras para tu búsqueda.');" & Chr(13))
                sb.Append("document." & strFormName & "." & strField & ".focus();" & Chr(13))
                sb.Append("document." & strFormName & "." & strField & ".select();" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("else {" & Chr(13))

                sb.Append("buttonclicks ++;" & Chr(13))

                sb.Append("}" & Chr(13)) 'end of first else

                sb.Append("if (buttonclicks == 1){" & Chr(13))
                sb.Append("document." & strFormName & ".btn" & strFormName & ".value = 'Espera...';" & Chr(13))
                sb.Append("document." & strFormName & ".btn" & strFormName & ".disabled = true;" & Chr(13))
                sb.Append("document." & strFormName & ".submit();" & Chr(13))
                sb.Append("}" & Chr(13)) 'end of second if

                'end verify form function
                sb.Append("}" & Chr(13))

                sb.Append("//-->" & Chr(13))
                sb.Append("</script>" & Chr(13))


                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function AdvancedSearch() As String

            Dim mForm As New ParaLideres.FormControls.GenericForm("frmSearch")
            Dim msb As StringBuilder = New StringBuilder("")

            Try

                msb.Append(mForm.FormAction("search.aspx", 450))

                'Param
                msb.Append(mForm.FormTextBox("Buscar Por", "shearchparam", "", 50, "Escribe aqu&#237; la palabra o frase para la b&#250;squeda", True))

                'Section
                msb.Append(mForm.FormSelect("Solo En", "section_id", 0, "TODO:", "Selecciona una secci&#243;n espec&#237;fica", False, "Todas Las Secciones"))

                'Book
                msb.Append(mForm.FormSelect("Libro de la Biblia", "book", 0, "sp_GetAllBooks", "Selecciona un libro de la Biblia", False))

                'Chapter
                msb.Append(mForm.FormTextBox("Cap&#237;tulo", "chapter", "", 3, "Si seleccionaste un libro de la Biblia puedes tambi&#233;n especificar un cap&#237;tulo dentro del libro", False, True))

                msb.Append(mForm.FormEnd("Buscar"))

                Return msb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                mForm = Nothing
                msb = Nothing

            End Try


        End Function

        Public Function HomePage() As String

            Try

                If IsNothing(Cache.Get("HomePageContent")) Then

                    Trace.Write("Generate home page content as put it on cache")

                    Dim sb As StringBuilder = New StringBuilder("")

                    Dim objArticle As DataLayer.pages

                    Dim intPageId As Integer = 0

                    Try

                        sb.Append("<table width=530 cellpadding=0 cellspacing=2 border=0>")

                        'TOP
                        intPageId = System.Web.Configuration.WebConfigurationManager.AppSettings.Get("HomeTop")

                        Trace.Write("intPageId: " & intPageId)

                        objArticle = New DataLayer.pages(intPageId)

                        sb.Append("<tr><td colspan=2 class=Estilo1>" & objArticle.getBody & "<br><br></td></tr>")

                        sb.Append("      <tr>")
                        sb.Append("        <td colspan=2 align=""left"" background=""" & _project_path & "" & "_images/puntos_horizontales.gif"" ><img src=""" & _project_path & "_images/puntos_horizontales.gif"" width=""1"" height=""1""></td>")
                        sb.Append("      </tr>")

                        sb.Append("<tr>")

                        'BOTTOM LEFT
                        intPageId = System.Web.Configuration.WebConfigurationManager.AppSettings.Get("HomeLeft")

                        Trace.Write("intPageId: " & intPageId)

                        objArticle = New DataLayer.pages(intPageId)
                        sb.Append("<td class=Estilo1>" & objArticle.getBody & "</td>")

                        'BOTTOM RIGHT
                        intPageId = System.Web.Configuration.WebConfigurationManager.AppSettings.Get("HomeRight")

                        Trace.Write("intPageId: " & intPageId)

                        objArticle = New DataLayer.pages(intPageId)
                        sb.Append("<td class=Estilo1>" & objArticle.getBody & "</td>")

                        sb.Append("</tr>")

                        sb.Append("</table>")

                        Cache.Insert("HomePageContent", sb.ToString())

                    Catch ex As Exception

                        ShowError(ex)

                    Finally

                        sb = Nothing

                        objArticle = Nothing

                    End Try

                End If

                PageTitle = "Bienvenido"
                PageContent = Cache.Get("HomePageContent")

            Catch ex As Exception

                ShowError(ex)

            End Try

        End Function

        Public Function PageTemplateTest() As String 'This is the current version used

            Dim sb As StringBuilder = New StringBuilder("")
            Dim strFlashFile As String = ""

            Try

                sb.Append("<html>")
                sb.Append("<head>")

                sb.Append("<title>" & _page_title & "</title>")

                sb.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">")

                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)

                sb.Append("</head>")

                sb.Append("<p>" & _page_title & "</p>")

                sb.Append("<p>" & _page_content & "</p>")

                sb.Append("<body>")

                sb.Append("</body>")
                sb.Append("</html>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function PageTemplate2() As String 'This is the current version used

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strFlashFile As String = ""

            Try

                sb.Append("<html>")
                sb.Append("<head>")

                sb.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">")

                sb.Append("<meta name=""verify-v1"" content=""lemJ1rsZeV1Q0103quVVmRiXpkvuQIw1Ja6JIF2sSf8="" />")

                sb.Append("<script language=""javascript"" src=""" & _project_path & "ajax.js"" type=""text/javascript""></script>" & vbLf)

                sb.Append("<link rel=""stylesheet"" href=""" & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""stylesheet"" href=""" & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)

                'RSS FEEDS
                sb.Append("<link rel=""alternate"" type=""application/rss+xml"" title=""Lo Ultimo en Para L&#237;deres"" href=""http://www.paralideres.org/rssfeed.aspx?rss_feed_id=1"" />" & vbLf)
                sb.Append("<link rel=""alternate"" type=""application/rss+xml"" title=""Noticias de Para L&#237;deres"" href=""http://www.paralideres.org/rssfeed.aspx?rss_feed_id=2"" />" & vbLf)
                sb.Append("<link rel=""alternate"" type=""application/rss+xml"" title=""Lo M&#225;s Destacado de Para L&#237;deres"" href=""http://www.paralideres.org/rssfeed.aspx?rss_feed_id=3"" />" & vbLf)
                sb.Append("<link rel=""alternate"" type=""application/rss+xml"" title=""Lo M&#225;s Visto de Para L&#237;deres"" href=""http://www.paralideres.org/rssfeed.aspx?rss_feed_id=4"" />" & vbLf)
                'sb.Append("<link rel=""alternate"" type=""application/rss+xml"" title=""Los Blogs de Para L&#237;deres"" href=""http://www.paralideres.org/rssfeed.aspx?rss_feed_id=5"" />" & vbLf)

                sb.Append("<title>" & _objUser.getFirstName & " - " & _page_title & "</title>")

                'sb.Append(Rollovers(_arrItems))

                sb.Append("</head>")

                sb.Append("<body ")

                'sb.Append(" onLoad=""MM_preloadImages('_images/enviar002.gif','_images/lo_ultimo02.gif','_images/cursos02.gif','_images/foros02.gif','_images/destacado02.gif','_images/preguntas_frecuentes02.gif')"" ")

                If _strOnLoadScript <> "" Then sb.Append(" onLoad=""" & _strOnLoadScript & """")

                sb.Append(">")


                If _simple_page Then

                    If Request("close") = "y" Then intDataGridWidth = 500 Else intDataGridWidth = 640

                    sb.Append("<p align=center valign=top>")

                    sb.Append("<table width=" & intDataGridWidth & " cellpadding=2 cellspacing=0 bordercolor=""#D09000"" border=""0"" style=""border-color:#D09000;border-width:1px;border-style:solid;width:" & intDataGridWidth & "px;border-collapse:collapse;background-color:white;"">")

                    sb.Append("<tr><td align=center valign=top>")

                    sb.Append("<table width=" & intDataGridWidth & " border=0 cellpadding=2 cellspacing=0>")

                    If Request("close") = "y" Then

                        sb.Append("<tr><td align=right width=""" & intDataGridWidth & """ height=15 bgcolor=""#A7D168""><a href=javascript:ClearAjaxContent();><img src=" & ProjectPath & "images/x.gif border=0 alt=""Cerrar""></a></td></tr>")

                        Me.Trace.IsEnabled = False

                    End If

                    'sbContent.Append("<tr><td width=""500"" bgcolor=" & Functions.HeaderBgColor & " align=right><a href=javascript:ClearAjaxContent();><img src=" & ProjectPath & "images/editor/x.gif border=0 alt=Close></a></td></tr>")

                    sb.Append("<tr><td align=left valign=top class=Estilo1>")

                    If _page_path <> "" Then sb.Append("<p align=left><b>" & _page_path & "</b></p>")

                    sb.Append("<p align=center><font size=3><b>" & _page_title & "</b></font></p>")

                    sb.Append("<p>")

                    sb.Append(_page_content)

                    sb.Append("</p>")

                    sb.Append("</td>")

                    sb.Append("</tr>")

                    'sb.Append("<tr><td align=center class=Estilo1><br><font size=3>" & _page_title & "</font><br><br></td></tr>")
                    'sb.Append("<tr><td valign=top class=Estilo1>" & _page_content & "</td></tr>")

                    'Bottom Menu
                    'sb.Append("<tr><td valign=top>")
                    'sb.Append(BottomMenu(530))
                    'sb.Append("</td></tr>")

                    sb.Append("</table>")

                    sb.Append("</td></tr></table></p>")

                Else

                    sb.Append("<div name=""divInstructions"" id=""divInstructions"" class='hideError' style=""FONT-FAMILY:Verdana;FONT-SIZE:9pt;BACKGROUND-COLOR:lightgoldenrodyellow;""></div>" & Chr(13))

                    sb.Append("<div name='divPanel' id='divPanel' align=center valign=top class=panel>" & Chr(13))

                    'sb.Append("<script language=""JavaScript"">" & Chr(13))

                    'sb.Append("document.write(""<div align=center valign=top name='divAjaxContent' id='divAjaxContent' onmouseover='over=true;' onmouseout='over=false;' class=AjaxDivContent></div>"")" & Chr(13))

                    sb.Append("<div align=center valign=top name='divAjaxContent' id='divAjaxContent' class=AjaxDivContent></div>" & Chr(13))

                    'sb.Append("</script>" & Chr(13))

                    sb.Append("</div>" & Chr(13))

                    sb.Append("<table width=""690"" height=""700"" border=""0"" valign=top align=center cellpadding=""0"" cellspacing=""0"">")

                    sb.Append("<tr>")

                    sb.Append("<td width=" & _intWidth & " align=""center"" valign=top style=""border-color:" & _strBorderColor & ";border-width:1px;border-style:solid;width:" & _intWidth & "px;"">")


                    '---------------------------------------------------------------------------------------
                    'Left Menu 
                    'Menu that has options on the left side of the page
                    If IsNothing(Cache.Get("LeftMenu")) Then

                        Dim obj As ParaLideres.MenuGenerator

                        Try

                            obj = New ParaLideres.MenuGenerator

                            obj.MaxLevel = 6 'depth of menu
                            obj.Width = _intWidth
                            obj.CellPadding = 3
                            obj.CellSpacing = 0
                            obj.BackColor = "White"
                            obj.ForeColor = "Black"
                            obj.BorderColor = "#D09000"
                            obj.HoverColor = "#A7D168"

                            Cache.Insert("LeftMenu", obj.GenerateMenu())

                            Trace.Write("GenerateMenu callled")

                        Catch ex As Exception

                            Trace.Warn(ex.ToString())

                        Finally

                            obj = Nothing

                        End Try

                    End If

                    sb.Append(Cache.Get("LeftMenu"))

                    'End of Left Menu
                    '---------------------------------------------------------------------------------------



                    '---------------------------------------------------------------------------------------
                    'Small Logon Form
                    'do not move this because it will not work on logon.aspx
                    'If _strScript <> "logon.aspx" And _objUser.getId = 0 Then

                    If _objUser.getId = 0 And _strScript <> "logon.aspx" Then

                        If IsNothing(Cache.Get("LogonForm")) Then

                            Cache.Insert("LogonForm", GenerateLeftTable("REG&#237;STRATE", SmallLogonForm()))

                            Trace.Write("LogonForm callled")

                        End If

                        sb.Append(Cache.Get("LogonForm"))

                    ElseIf _objUser.getId > 0 Then

                        If IsNothing(Cache.Get("MiParaLideres")) Then

                            Dim sb2 As New System.Text.StringBuilder("")

                            Try

                                sb2.Append("<tr>" & Chr(13))

                                sb2.Append("<td  class=""Estilo1"" align=left>")

                                sb2.Append("<img src=""_images/item.gif""><a href=registration.aspx>Ver Mis Datos</a><br><br>")

                                sb2.Append("<img src=""_images/item.gif""><a href=pages_by_user.aspx?user_id=0>Ver Mis Publicaciones</a><br><br>")

                                sb2.Append("<img src=""_images/item.gif""><a href=mis_favoritos.aspx>Ver Mis Art&#237;culos Favoritos</a><br><br>")

                                sb2.Append("<img src=""_images/item.gif""><a href=publish_my_materials.aspx>Colaborar Con Una Nueva Publicaci&#243;n</a><br><br>")

                                sb2.Append("<img src=""_images/item.gif""><a href=logoff.aspx>Desconectarse</a><br><br>")

                                sb2.Append("</td>" & Chr(13))

                                sb2.Append("</tr>" & Chr(13))

                                Cache.Insert("MiParaLideres", GenerateLeftTable("MI PARA L&#237;DERES", sb2.ToString()))

                            Catch ex As Exception

                                Throw ex

                            Finally

                                sb2 = Nothing

                            End Try

                        End If

                        sb.Append(Cache.Get("MiParaLideres"))

                    End If
                    '-------------------------------------------------------------------------------


                    '---------------------------------------------------------------------------------------
                    'Survey
                    sb.Append(GenerateLeftTable("¿Qu&#233; Piensas?", DisplaySurveyInfo()))
                    'End Survey
                    '---------------------------------------------------------------------------------------

                    '--------------------------------------------------------------------------------------
                    '--Add
                    '--------------------------------------------------------------------------------------
                    If IsNothing(Cache.Get("PageAdd")) Then

                        Dim sb1 As StringBuilder

                        Dim strLink As String = "http://www.institutoej.com"
                        Dim strImageName As String = "iej.jpg"

                        Try

                            sb1 = New StringBuilder("")

                            sb1.Append("<tr>" & Chr(13))

                            sb1.Append("<td  class=""Estilo1"" align=""center"">")

                            sb1.Append("<a href=""" & strLink & """ target=new><img src=" & ProjectPath & "images/" & strImageName & " border=0></a>")

                            sb1.Append("</td>" & Chr(13))

                            sb1.Append("</tr>" & Chr(13))

                            Cache.Insert("PageAdd", GenerateLeftTable(" ", sb1.ToString()))

                        Catch ex As Exception

                            Throw ex

                        Finally

                            sb1 = Nothing

                        End Try

                    End If

                    sb.Append(Cache.Get("PageAdd"))

                    '-------------------------------------------------------------------------------





                    '-------------------------------------------------------------------------------
                    '---Articles last viewed
                    '-------------------------------------------------------------------------------
                    If IsNothing(Cache.Get("LastViewed")) Then

                        Cache.Insert("LastViewed", _arrLastViewed)

                        Cache.Insert("LastViewedIndex", 0)

                    Else

                        Dim sb1 As StringBuilder

                        Try

                            sb1 = New StringBuilder("")

                            _arrLastViewed = Cache.Get("LastViewed")

                            If _arrLastViewed.GetUpperBound(0) > 0 Then

                                sb1.Append("<tr>" & Chr(13))

                                sb1.Append("<td  class=""Estilo1"" align=left>")

                                Dim intX As Integer = 0

                                For intX = 0 To _arrLastViewed.GetUpperBound(0)

                                    sb1.Append(_arrLastViewed(intX, 0))

                                Next

                                sb1.Append("</td>" & Chr(13))

                                sb1.Append("</tr>" & Chr(13))

                                sb.Append(GenerateLeftTable("RECIENTEMENTE VISTO", sb1.ToString()))

                            End If

                        Catch ex As Exception

                            Throw ex

                        Finally

                            sb1 = Nothing

                        End Try

                    End If
                    '-------------------------------------------------------------------------------

                    sb.Append("<br>" & DisplayFlag())

                    sb.Append("</td>")
                    'end of left cell 


                    'right cell
                    sb.Append("<td width=530 align=""left"" valign=top style=""border-color:" & _strBorderColor & ";border-width:1px;border-style:solid;width:530px;"">")

                    sb.Append("     <table width=""530"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")

                    'If _strScript = "home.aspx" Then

                    '    sb.Append("<tr>" & Chr(13))
                    '    sb.Append("<td><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""530"" height=""130"">" & Chr(13))
                    '    sb.Append("<param name=""movie"" value=""header.swf"">" & Chr(13))
                    '    sb.Append("<param name=""quality"" value=""high"">" & Chr(13))
                    '    sb.Append("<embed src=""header.swf"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""530"" height=""130""></embed>" & Chr(13))
                    '    sb.Append("</object></td>" & Chr(13))
                    '    sb.Append("</tr>" & Chr(13))

                    'Else

                    sb.Append("      <tr>")
                    sb.Append("        <td><a href=" & _project_path & "home.aspx><img src=""" & _project_path & "_images/header_contenido.jpg"" width=""530"" height=""67"" border=0></a></td>")
                    sb.Append("      </tr>")


                    'End If

                    sb.Append("      <tr>")
                    sb.Append("        <td height=""6""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""10""></td>")
                    sb.Append("      </tr>")

                    'search form

                    sb.Append(Search())

                    sb.Append("      <tr>")
                    sb.Append("        <td align=""left"" background=""" & _project_path & "" & "_images/puntos_horizontales.gif""><img src=""" & _project_path & "_images/puntos_horizontales.gif"" width=""3"" height=""1""></td>")
                    sb.Append("      </tr>")


                    sb.Append("      <tr>")
                    sb.Append("        <td height=""4""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""4""></td>")
                    sb.Append("      </tr>")

                    'Estas Aqui/You are here

                    If _page_path <> "" Then

                        sb.Append("      <tr>")
                        sb.Append("        <td class=DIM>")
                        sb.Append("             Est&#225;s aqu&#237; >> " & _page_path & "<br><br>")
                        sb.Append("        </td>")
                        sb.Append("     </tr>")

                    End If


                    sb.Append("      <tr>")
                    sb.Append("        <td height=""15""><img src=""" & _project_path & "_images/barra_larga.gif"" width=""530"" height=""15""></td>")
                    sb.Append("      </tr>")

                    sb.Append("    </table>")

                    sb.Append("          <table width=""527"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                    sb.Append("            <tr>")
                    sb.Append("              <td width=""515"" height=""48"" class=""Estilo1"" valign=top>")

                    '<div align=""left"" valign=top>")

                    '---------------------------------------------------------------------------------------
                    sb.Append("                <p align=center><font size=3><b>" & _page_title & "</b></font></p>")
                    '---------------------------------------------------------------------------------------

                    sb.Append("<p>")

                    '---------------------------------------------------------------------------------------
                    sb.Append(_page_content)
                    '---------------------------------------------------------------------------------------

                    sb.Append("                 </p>")

                    If Request("send_email") = "yes" Then

                        sb.Append(EmailPageForm())

                    End If

                    'sb.Append("</div>")

                    sb.Append("              </td>")
                    sb.Append("            </tr>")
                    sb.Append("          </table>")

                    sb.Append(" </td>")
                    'end of right cell

                    sb.Append("  </tr>")
                    'end of second row


                    'bottom menu row
                    sb.Append("  <tr>")
                    sb.Append("    <td colspan=""2"">")

                    '---------------------------------------------------------------------------------------
                    'Bottom Menu
                    sb.Append(BottomMenu())
                    '---------------------------------------------------------------------------------------
                    sb.Append("    </td>")
                    sb.Append("  </tr>")
                    'end of bottom menu row


                    sb.Append("</table>")

                End If

                sb.Append("</body>")
                sb.Append("</html>")

                Return sb.ToString()


            Catch ex As Exception

                Functions.ShowError(ex)

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function

        Private Function EmailPageForm() As String

            Session("emailcontent") = _page_content
            Session("emailsubject") = _page_title

            Dim sb As New System.Text.StringBuilder("")
            Dim objFrm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("frmContact")

            Try
                sb.Append("<a name=""email""></a>")
                sb.Append(objFrm.FormAction("email_me.aspx"))
                sb.Append(objFrm.FormEmailVerifyBox("Enviar a", "recipient", "", "Ingresa el e-mail a quien vas a enviar este art&#237;culo"))
                sb.Append(objFrm.FormTextArea("Comentarios", "comments", "Mira este art&#237;culo publicado en ParaL&#237;deres.org", 10, 50, "Ingresa comentarios que quieras mandar.", False))
                sb.Append(objFrm.FormEnd("Enviar E-mail"))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                objFrm = Nothing

            End Try

        End Function

        Public Function PageTemplateHtml(ByVal strDescription As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                _simple_page = True

                sb.Append("<html>")
                sb.Append("<head>")

                sb.Append("<title>" & _page_title & "</title>")

                sb.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">")

                sb.Append("<meta name=""robots"" content=""index,nofollow"">")
                sb.Append("<meta name=""description"" content=""" & strDescription & """>")

                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)


                'sb.Append(Rollovers(_arrItems))

                sb.Append("</head>")

                sb.Append("<body ")

                'sb.Append(" onLoad=""MM_preloadImages('_images/enviar002.gif','_images/lo_ultimo02.gif','_images/cursos02.gif','_images/foros02.gif','_images/destacado02.gif','_images/preguntas_frecuentes02.gif')"" ")

                If _strOnLoadScript <> "" Then sb.Append(" onLoad=""" & _strOnLoadScript & """")


                sb.Append(">")


                If _simple_page Then

                    sb.Append("<p align=center valign=top>")
                    sb.Append("<table width=530 cellpadding=2 cellspacing=0 bordercolor=""#D09000"" border=""1"" style=""border-color:#D09000;border-width:1px;border-style:solid;width:530px;border-collapse:collapse;"">")

                    sb.Append("<tr><td align=center valign=top>")

                    sb.Append("<table border=0 cellpadding=0 cellspacing=0>")
                    sb.Append("<tr><td align=center><img src=" & _project_path & "_images/header_contenido.jpg border=0><br></td></tr>")

                    sb.Append("<tr><td align=center class=Estilo1>")

                    sb.Append("                <p align=center><font size=3><b>" & _page_title & "</b></font></p>")

                    sb.Append("                <p>")

                    sb.Append(_page_content)

                    sb.Append("                 </p>")

                    sb.Append("                 </td>")
                    sb.Append("                 </tr>")

                    'sb.Append("<tr><td align=center class=Estilo1><br><font size=3>" & _page_title & "</font><br><br></td></tr>")
                    'sb.Append("<tr><td valign=top class=Estilo1>" & _page_content & "</td></tr>")

                    'Bottom Menu
                    'sb.Append("<tr><td valign=top>")
                    'sb.Append(BottomMenu(530))
                    'sb.Append("</td></tr>")

                    sb.Append("</table>")

                    sb.Append("</td></tr></table></p>")

                Else


                    sb.Append("<table width=""690"" border=""0"" valign=top align=center cellpadding=""0"" cellspacing=""0"">")

                    'first row
                    sb.Append("<tr>")

                    'left cell
                    sb.Append("<td bgcolor=white width=146 height=45 valign=bottom>")
                    sb.Append("<img src=""" & _project_path & "_images/separador_45px.gif"" width=""146"" height=""30"">")
                    sb.Append("<img src=""" & _project_path & "_images/desplegable_image.jpg"" width=""146"" height=""15"">")
                    sb.Append("</td>")
                    'end left cell

                    'right cell 'Flash Menu
                    sb.Append("    <td align=""left"" valign=""top"">")
                    'sb.Append("     <table width=""530"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
                    'sb.Append("      <tr>")
                    'sb.Append("        <td valign=""top""><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""530"" height=""45"">")
                    'sb.Append("          <param name=""movie"" value=""menu_contenido.swf"">")
                    'sb.Append("          <param name=""quality"" value=""high"">")
                    'sb.Append("          <embed src=""menu_contenido.swf"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""530"" height=""45""></embed>")
                    'sb.Append("        </object></td>")
                    'sb.Append("      </tr>")
                    'sb.Append("     </table>")
                    sb.Append("    </td>")
                    'end right cell

                    'end of first row
                    sb.Append("</tr>")
                    '---------------------------------------------------------------------------------------


                    'second row
                    sb.Append("<tr>")

                    'left cell second row
                    sb.Append("<td width=146 align=""center"" valign=top background=""" & _project_path & "_images/fondo_celda_izquierda.gif"">")

                    sb.Append("<table width=""140"" border=""0"" align=center valign=top cellspacing=""0"" cellpadding=""0"">")

                    '---------------------------------------------------------------------------------------
                    'drop down menu

                    If IsNothing(Cache.Get("DropDown")) Then

                        Cache.Insert("DropDown", DropDownMenu())

                        Trace.Write("DropDownMenu callled")

                    End If

                    sb.Append(Cache.Get("DropDown"))

                    '---------------------------------------------------------------------------------------

                    ''---------------------------------------------------------------------------------------
                    ''Left Menu 
                    ''Menu that has options on the left side of the page
                    'If IsNothing(Cache.Get("LeftMenu")) Then

                    '    Cache.Insert("LeftMenu", GenerateMenu())

                    '    Trace.Write("GenerateMenu callled")

                    'End If

                    'sb.Append(Cache.Get("LeftMenu"))
                    ''End of Left Menu
                    ''End of Left Menu
                    '---------------------------------------------------------------------------------------

                    ''---------------------------------------------------------------------------------------
                    ''Small Logon Form
                    'If _strScript = "home.aspx" And _objUser.getId = 0 Then

                    '    sb.Append(Me.SmallLogonForm())

                    'End If
                    '---------------------------------------------------------------------------------------


                    '---------------------------------------------------------------------------------------
                    'Survey
                    'sb.Append(DisplaySurveyInfo())
                    'End Survey
                    '---------------------------------------------------------------------------------------


                    'If _objUser.getId > 0 Then


                    '    sb.Append("<tr>" & Chr(13))
                    '    sb.Append("<td><br><br></td>" & Chr(13))
                    '    sb.Append("</tr>" & Chr(13))

                    '    sb.Append("<tr>" & Chr(13))
                    '    sb.Append("<td><img src=""_images/registrate_colabora.gif"" width=""146"" height=""15""></td>" & Chr(13))
                    '    sb.Append("</tr>" & Chr(13))

                    '    sb.Append("<tr>" & Chr(13))
                    '    sb.Append("<td  class=""Estilo1"">Usuario: " & Chr(13))
                    '    sb.Append(_objUser.getFirstName & " " & _objUser.getLastName)
                    '    sb.Append("</td>" & Chr(13))
                    '    sb.Append("</tr>" & Chr(13))

                    'End If

                    sb.Append("</table>")

                    sb.Append("</td>")
                    'end of left cell second row


                    'right cell
                    sb.Append("    <td align=""left"" valign=""top"">")
                    sb.Append("     <table width=""530"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")

                    'If _strScript = "home.aspx" Then

                    '    sb.Append("<tr>" & Chr(13))
                    '    sb.Append("<td><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""530"" height=""130"">" & Chr(13))
                    '    sb.Append("<param name=""movie"" value=""header.swf"">" & Chr(13))
                    '    sb.Append("<param name=""quality"" value=""high"">" & Chr(13))
                    '    sb.Append("<embed src=""header.swf"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""530"" height=""130""></embed>" & Chr(13))
                    '    sb.Append("</object></td>" & Chr(13))
                    '    sb.Append("</tr>" & Chr(13))

                    'Else

                    sb.Append("      <tr>")
                    sb.Append("        <td><a href=" & _project_path & "home.aspx><img src=""" & _project_path & "_images/header_contenido.jpg"" width=""530"" height=""67"" border=0></a></td>")
                    sb.Append("      </tr>")


                    'End If

                    sb.Append("      <tr>")
                    sb.Append("        <td height=""6""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""10""></td>")
                    sb.Append("      </tr>")

                    'search form

                    'sb.Append(Search())

                    sb.Append("      <tr>")
                    sb.Append("        <td align=""left"" background=""" & _project_path & "" & _project_path & "_images/puntos_horizontales.gif""><img src=""" & _project_path & "_images/puntos_horizontales.gif"" width=""3"" height=""1""></td>")
                    sb.Append("      </tr>")
                    sb.Append("      <tr>")
                    sb.Append("        <td height=""4""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""4""></td>")
                    sb.Append("      </tr>")

                    'Estas Aqui/You are here
                    sb.Append("      <tr>")

                    sb.Append("        <td height=""15"" class=DIM>")

                    sb.Append("           <img src=""" & _project_path & "_images/puntos_horizontales.gif"" >")

                    If _page_path <> "" Then sb.Append("<br>Est&#225;s aqu&#237; >> " & _page_path)

                    sb.Append("        </td>")

                    sb.Append("     </tr>")


                    sb.Append("      <tr>")
                    sb.Append("        <td height=""15""><img src=""" & _project_path & "_images/barra_larga.gif"" width=""530"" height=""15""></td>")
                    sb.Append("      </tr>")

                    sb.Append("    </table>")

                    sb.Append("          <table width=""527"" border=""0"" cellspacing=""6"" cellpadding=""0"">")
                    sb.Append("            <tr>")
                    sb.Append("              <td width=""515"" height=""48"" class=""Estilo1"" valign=top><div align=""left"" valign=top>")

                    '---------------------------------------------------------------------------------------
                    sb.Append("                <p align=center><font size=3><b>" & _page_title & "</b></font></p>")
                    '---------------------------------------------------------------------------------------

                    sb.Append("<p>")

                    '---------------------------------------------------------------------------------------
                    sb.Append(_page_content)
                    '---------------------------------------------------------------------------------------

                    sb.Append("                 </p>")
                    sb.Append("              </div></td>")
                    sb.Append("            </tr>")
                    sb.Append("          </table>")

                    sb.Append(" </td>")
                    'end of right cell

                    sb.Append("  </tr>")
                    'end of second row


                    'bottom menu row
                    sb.Append("  <tr>")
                    sb.Append("    <td colspan=""2"">")

                    '---------------------------------------------------------------------------------------
                    'Bottom Menu
                    sb.Append(BottomMenu())
                    '---------------------------------------------------------------------------------------
                    sb.Append("    </td>")
                    sb.Append("  </tr>")
                    'end of bottom menu row


                    sb.Append("</table>")

                End If

                sb.Append("</body>")
                sb.Append("</html>")

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function BottomMenu(Optional ByVal intMenuWidth As Integer = 700) As String

            If IsNothing(Cache.Get("BottomMenu")) Then

                Dim sb As New StringBuilder("")

                Try

                    sb.Append("    <table width=""" & intMenuWidth & """ border=""0"" cellspacing=""0"" cellpadding=""0"">")
                    sb.Append("      <tr>")
                    sb.Append("        <td height=""6""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""6""></td>")
                    sb.Append("      </tr>")
                    sb.Append("      <tr>")
                    sb.Append("        <td bgcolor=""#CC3300""><img src=""" & _project_path & "_images/abajo_rojo.gif"" width=""6"" height=""6""></td>")
                    sb.Append("      </tr>")
                    sb.Append("      <tr>")
                    sb.Append("        <td height=""2""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""2""></td>")
                    sb.Append("      </tr>")
                    sb.Append("      <tr>")
                    sb.Append("        <td height=""20"" bgcolor=""#FF9900"" class=""Estilo2"" align=center valign=middle>")

                    'sb.Append("         preguntas frecuentes <span class=""Estilo3"">I</span>")

                    sb.Append("         <a href=" & _project_path & "article.aspx?page_id=653 target=new class=blackbold>qui&eacute;n es ParaL&iacute;deres</a> <span class=""Estilo3"">I</span>")
                    sb.Append("         <a href=" & _project_path & "article.aspx?page_id=652 target=new class=blackbold>qu&eacute; creemos</a> <span class=""Estilo3"">I</span>")
                    sb.Append("         <a href=" & _project_path & "article.aspx?page_id=216 target=new class=blackbold>condiciones de servicio</a> <span class=""Estilo3"">I</span>")
                    sb.Append("         <a href=" & _project_path & "comments.aspx>e-mail</a>")

                    sb.Append("        </td>")

                    sb.Append("      </tr>")
                    sb.Append("    </table>")


                    Cache.Insert("BottomMenu", sb.ToString)

                    Trace.Write("BottomMenu() callled")

                Catch ex As Exception

                    Throw ex

                Finally

                    sb = Nothing

                End Try


            End If

            Return Cache.Get("BottomMenu")

        End Function

        Public Function QueryStringUrl(Optional ByVal strParamater As String = "") As String

            Dim objItem As Object
            Dim sb As StringBuilder

            Try

                sb = New StringBuilder("")

                sb.Append(Me.Request.Path() & "?")

                For Each objItem In Request.QueryString

                    sb.Append(objItem & "=" & Request(objItem) & "&")

                Next


                For Each objItem In Request.Form

                    sb.Append(objItem & "=" & Request(objItem) & "&")

                Next

                sb.Replace(" ", "%20")

                If strParamater <> "" Then

                    sb.Append(strParamater)

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objItem = Nothing

                sb = Nothing

            End Try

        End Function



        Public Function ConfirmDelete(ByVal strElement As String, ByVal strUrl As String) As String

            Dim sb As New StringBuilder("")

            sb.Append(Chr(13) & "<script language=javascript>" & Chr(13))
            sb.Append("function ConfirmDelete" & strElement & "(ps_id, ps_description)" & Chr(13))
            sb.Append("{" & Chr(13))
            sb.Append("myvar = confirm('Warning! \n \n Are you sure you want to delete ' + ps_description + '?');" & Chr(13))
            sb.Append("if (myvar) {" & Chr(13))

            sb.Append("myURL = '" & strUrl & "' + ps_id; " & Chr(13))
            sb.Append("x = setTimeout('window.location.href=myURL',0); " & Chr(13))
            sb.Append("}" & Chr(13))
            sb.Append("}" & Chr(13))
            sb.Append("</script>" & Chr(13))

            Return sb.ToString

        End Function


        Private Function GenerateLeftTable(ByVal strTitle As String, ByVal strRowsWithContent As String) As String

            Dim sb As New System.Text.StringBuilder("")


            Try

                sb.Append(Chr(13) & "<!-- " & UCase(strTitle) & " -->" & Chr(13))

                sb.Append("<table cellpadding=2 cellspacing=1 style=""border-color:" & _strBorderColor & ";border-width:1px;border-style:solid;width:" & _intWidth & "px;border-collapse:collapse;"">" & Chr(13))

                sb.Append("<tr><td align=center style=""width:" & _intWidth & "px;background:" & _strBackGroundColor & ";color:white;text-decoration:none;font-family:Verdana;font-size:8pt;font-weight:bold;"">" & UCase(strTitle) & "</td></tr>" & Chr(13))

                sb.Append(strRowsWithContent)

                sb.Append("<tr><td height=""1"" bgcolor=#CCCCCC></td></tr>" & Chr(13))

                sb.Append("</table>" & Chr(13))

                sb.Append("<!-- END " & UCase(strTitle) & " -->" & Chr(13))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function

        Private Function DisplaySurveyInfo() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim objSurvey As New ParaLideres.Survey(Date.Today)

            Try

                'sb.Append("<tr>")
                'sb.Append("<td valign=top><img src=""" & _project_path & "_images/que_piensas.gif"" width=""146"" height=""15""></td>")
                'sb.Append("</tr>")

                sb.Append("<tr><td align=""center"" valign=""top"">")

                If Session("sessAlreadyVoted") Then

                    'sb.Append(objSurvey.DisplayCurrentSurveyResults())

                    sb.Append(objSurvey.DisplayCurrentQuestion())

                Else

                    sb.Append(objSurvey.DisplayCurrentSurvey())

                End If

                sb.Append("</td></tr>")

                Return sb.ToString()

            Catch ex As Exception

                Functions.ShowError(ex)

                Throw ex

            Finally

                sb = Nothing
                objSurvey = Nothing

            End Try

        End Function

        Private Function SmallLogonForm() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim strUsername As String = ""
            Dim strFldUsername As String = ""
            Dim strFldPassword As String = ""
            Dim strFormName As String = "frmLogon"
            Dim intRandom As Integer = ParaLideres.Functions.GetRandomNumber(500, 999)

            If Not IsNothing(Request.Cookies("username")) Then

                strUsername = Request.Cookies("username").Value

            End If

            'This is to prevent dictionary and brute force attaks
            strFldUsername = "username" & intRandom
            strFldPassword = "password" & intRandom

            Try

                'sb.Append("<tr>" & Chr(13))
                'sb.Append("<td><img src=""_images/registrate_colabora.gif"" width=""146"" height=""15""></td>" & Chr(13))
                'sb.Append("</tr>" & Chr(13))

                'Form Row
                sb.Append("<tr>" & Chr(13))
                sb.Append("<td align=""center"">" & Chr(13))

                sb.Append("<form name=""" & strFormName & """ method=""post"" action=""chcklog.aspx"">" & Chr(13))
                sb.Append("<input type=hidden name=url_redirect value=home.aspx>")

                sb.Append("<table width=""140"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & Chr(13))

                'username label
                sb.Append("<tr>" & Chr(13))
                sb.Append("<td><img src=""_images/email.gif"" width=""140"" height=""15""></td>" & Chr(13))
                sb.Append("</tr>" & Chr(13))

                'username fld
                sb.Append("<tr>" & Chr(13))
                sb.Append("<td><input type=""text"" name=""" & strFldUsername & """ value=""" & strUsername & """ Class=frmInputSmall></td>" & Chr(13))
                sb.Append("</tr>" & Chr(13))

                'password label
                sb.Append("<tr>" & Chr(13))
                sb.Append("<td><img src=""_images/clave.gif"" width=""140"" height=""15""></td>" & Chr(13))
                sb.Append("</tr>" & Chr(13))

                'password fld
                sb.Append("<tr>" & Chr(13))
                sb.Append("<td><input type=""password"" name=""" & strFldPassword & """ Class=frmInputSmall maxlength=16></td>" & Chr(13))
                sb.Append("</tr>" & Chr(13))

                'Button
                sb.Append("<tr>")
                sb.Append("<td colspan=2 align=center><input id=btn" & strFormName & " type=button value=""Entrar"" alt=""Entrar""  onclick='VerifyForm" & strFormName & "();' class=frmbutton  height=25></td>")
                sb.Append("</tr>")

                'Link to e-mail password
                'sb.Append("<tr>" & Chr(13))
                'sb.Append("<td><a href=emailpassword.aspx><img src=""_images/olvidaste_clave.gif"" width=""140"" height=""15"" border=""0""></a></td>" & Chr(13))
                'sb.Append("</tr>" & Chr(13))

                sb.Append("<tr>" & Chr(13))
                sb.Append("<td>")

                sb.Append("<p align=center><a href=emailpassword.aspx>¿Has olvidado tu clave?</a></p>" & Chr(13))
                sb.Append("<p align=center><a href=registration.aspx>Reg&#237;strate como nuevo usuario</a></p>" & Chr(13))

                sb.Append("</td>" & Chr(13))
                sb.Append("</tr>" & Chr(13))

                sb.Append("</table>" & Chr(13))

                sb.Append("</form>" & Chr(13))

                sb.Append("</td>" & Chr(13))

                sb.Append("</tr>" & Chr(13))

                'Form script
                sb.Append("<script language=javascript>" & Chr(13))
                sb.Append("<!--" & Chr(13))

                sb.Append("var whitespace = ' \t\n\r';" & Chr(13))

                'function VerifyForm()
                sb.Append(Chr(13) & "function VerifyForm" & strFormName & "(){" & Chr(13))

                sb.Append("var buttonclicks = 0;" & Chr(13))

                'check for username
                sb.Append("if (document." & strFormName & "." & strFldUsername & ".value == """"){" & Chr(13))
                sb.Append("	alert('Por favor ingresa tu e-mail');" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldUsername & ".focus();" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldUsername & ".select();" & Chr(13))
                sb.Append("}" & Chr(13))


                sb.Append("//email input validation for E-mail" & Chr(13))
                sb.Append("else if (!isEmail(document." & strFormName & "." & strFldUsername & ".value)){" & Chr(13))
                sb.Append("alert('Por favor ingresa una direcci&#243;n de e-mail que sea v&#225;lida');" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldUsername & ".focus();" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldUsername & ".select();" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("else if (document." & strFormName & "." & strFldPassword & ".value == """"){" & Chr(13))
                sb.Append("	alert('Por favor ingresa tu clave');" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldPassword & ".focus();" & Chr(13))
                sb.Append("document." & strFormName & "." & strFldPassword & ".select();" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("else {" & Chr(13))

                sb.Append("buttonclicks ++;" & Chr(13))

                sb.Append("}" & Chr(13)) 'end of first else

                sb.Append("if (buttonclicks == 1){" & Chr(13))
                sb.Append("document." & strFormName & ".btn" & strFormName & ".value = 'Espera...';" & Chr(13))
                sb.Append("document." & strFormName & ".btn" & strFormName & ".disabled = true;" & Chr(13))
                sb.Append("document." & strFormName & ".submit();" & Chr(13))
                sb.Append("}" & Chr(13)) 'end of second if

                'end verify form function
                sb.Append("}" & Chr(13))


                'function isEmpty
                sb.Append("function isEmpty(s){" & Chr(13))
                sb.Append("   return ((s == null) || (s.length == 0));" & Chr(13))
                sb.Append("}" & Chr(13))
                'end isEmpty

                'function isWhitespace()
                sb.Append("function isWhitespace (s){" & Chr(13))
                sb.Append("    var i;" & Chr(13))
                sb.Append("    if (isEmpty(s)) return true;" & Chr(13))
                sb.Append("    for (i = 0; i < s.length; i++) {" & Chr(13))
                sb.Append("        var c = s.charAt(i);" & Chr(13))
                sb.Append("        if (whitespace.indexOf(c) == -1) return false;" & Chr(13))
                sb.Append("    }" & Chr(13))
                sb.Append("    return true;" & Chr(13))
                sb.Append("}" & Chr(13))
                'end isWhitespace

                'validate e-mail address

                sb.Append("function isEmail (s){" & Chr(13))
                sb.Append("if (isWhitespace(s)) return false;" & Chr(13))
                sb.Append("var i = 1;" & Chr(13))
                sb.Append("var sLength = s.length;" & Chr(13))
                sb.Append("while ((i < sLength) && (s.charAt(i) != '@')){" & Chr(13))
                sb.Append("   i++" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("if ((i >= sLength) || (s.charAt(i) != '@')) return false;" & Chr(13))
                sb.Append("else i += 2;" & Chr(13))

                sb.Append("while ((i < sLength) && (s.charAt(i) != '.')){" & Chr(13))
                sb.Append("   i++" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("if ((i >= sLength - 1) || (s.charAt(i) != '.')) return false;" & Chr(13))

                sb.Append("else return true;" & Chr(13))
                sb.Append("}" & Chr(13))

                'end isEmail()

                sb.Append("//-->" & Chr(13))
                sb.Append("</script>" & Chr(13))


                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function

        Public Function CommentsForm() As String

            Dim mForm As New ParaLideres.FormControls.GenericForm("frmComments")
            Dim msb As StringBuilder = New StringBuilder("")

            Try

                msb.Append("Es muy importante para nosotros saber tus comentarios y/o sugerencias para as&#237; poderte servirte mejor.")

                msb.Append(mForm.FormAction("comments.aspx"))

                'Email
                msb.Append(mForm.FormTextBox("E-mail", "email", _objUser.getEmail, 50, "Inserte su direcci&#243;n de E-mail", True, 5, False, True))

                'Comments
                msb.Append(mForm.FormTextArea("Comentarios/Sugerencias", "comments", "", 7, 50, "Ingresa los comentarios o sugerencias que quieras compartir con nosotros.", True, 400))

                msb.Append(mForm.FormTextBox("N&#250;mero", "captcha", "", 5, "Por favor escribe el n&#250;mero que se muestra en la imagen.", True, 5, True, False, True))

                msb.Append(mForm.FormEnd())

                Return msb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                mForm = Nothing
                msb = Nothing

            End Try

        End Function



        Private Function DisplayCurrentSurvey() As String

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As New StringBuilder("")
            Dim intQuestionID As Integer = 0
            Dim strQuestionDesc As String = ""
            Dim arrQuestions() As String
            Dim intIndex As Integer = 0
            Dim strThisQuestion As String = ""


            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("sp_GetCurrentSurvey '" & Date.Today() & "'")

                If reader.HasRows Then

                    Do While reader.Read()

                        intQuestionID = reader(0)
                        strQuestionDesc = reader(1)

                        For intIndex = 4 To 13

                            strThisQuestion = strThisQuestion & reader(intIndex) & "|"

                        Next

                    Loop

                End If

                reader.Close()

                If strThisQuestion.Length > 0 Then

                    strThisQuestion = Left(strThisQuestion, Len(strThisQuestion) - 1)

                    arrQuestions = Split(strThisQuestion, "|")

                    sb.Append("          <table width=""140"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                    sb.Append("            <tr>")
                    sb.Append("              <td height=""37"" colspan=""2""><span class=""Estilo1"">" & strQuestionDesc & "</span><br><br></td>")
                    sb.Append("              </tr>")


                    For intIndex = LBound(arrQuestions) To UBound(arrQuestions)

                        If Len(arrQuestions(intIndex)) > 0 Then

                            sb.Append("            <tr>")
                            sb.Append("              <td width=""22"" align=""left""><input type='radio' name='optvote' value='" & intIndex + 1 & "' ></td>")
                            sb.Append("              <td width=""118""><div align=""left""><span class=""Estilo1"">" & arrQuestions(intIndex) & "</span></div></td>")
                            sb.Append("            </tr>")

                        End If

                    Next

                    sb.Append("            <tr>")
                    sb.Append("              <td height=""40"" colspan=""2"" align=""left""><div align=""center"">")


                    sb.Append(GenerateMenuItem("survey", "post_survey.aspx"))

                    'sb.Append("              <a href=""#"" onMouseOut=""MM_swapImgRestore()"" onMouseOver=""MM_swapImage('Image20','','" & _project_path & "_images/enviar002.gif',1)""><img src=""" & _project_path & "_images/enviar001.gif"" name=""Image20"" width=""55"" height=""20"" border=""0""></a>")

                    sb.Append("              </div></td>")
                    sb.Append("              </tr>")
                    sb.Append("          </table>")

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                reader = Nothing
                sb = Nothing

            End Try



        End Function

        'Public Function DisplayCurrentSurveyResults() As String

        '    Dim reader As System.Data.SqlClient.SqlDataReader
        '    Dim sb As New StringBuilder("")
        '    Dim intQuestionID As Integer = 0
        '    Dim strQuestionDesc As String = ""
        '    Dim arrQuestions() As String
        '    Dim arrAnswers() As String
        '    Dim intIndex As Integer = 0
        '    Dim strThisQuestion As String = ""
        '    Dim strThisAnswer As String = ""
        '    Dim intTotalVotes As Integer = 0
        '    Dim dblPercentage As Double = 0


        '    Try

        '        reader =ParaLideres.GenericDataHandler.GetRecords("sp_GetCurrentSurvey '" & Date.Today() & "'")

        '        If reader.HasRows Then

        '            Do While reader.Read()

        '                intQuestionID = reader(0)
        '                strQuestionDesc = reader(1)

        '                For intIndex = 4 To 13

        '                    strThisQuestion = strThisQuestion & reader(intIndex) & "|"

        '                Next

        '            Loop

        '        End If

        '        reader.Close()

        '        If strThisQuestion.Length > 0 Then

        '            strThisQuestion = Left(strThisQuestion, Len(strThisQuestion) - 1)

        '            arrQuestions = Split(strThisQuestion, "|")


        '            'get results
        '            reader =ParaLideres.GenericDataHandler.GetRecords("sp_GetTotalAnswersByQuestionID " & intQuestionID)

        '            If reader.HasRows Then

        '                Do While reader.Read()

        '                    For intIndex = 0 To 9

        '                        strThisAnswer = strThisAnswer & reader(intIndex) & "|"
        '                        intTotalVotes = intTotalVotes + reader(intIndex)

        '                    Next

        '                Loop

        '            End If

        '            reader.Close()


        '            If strThisAnswer.Length > 0 Then


        '                strThisAnswer = Left(strThisAnswer, Len(strThisAnswer) - 1)

        '                arrAnswers = Split(strThisAnswer, "|")


        '                sb.Append("          <table width=""150"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
        '                sb.Append("            <tr>")
        '                sb.Append("              <td height=""37"" colspan=""2""><span class=""Estilo1""><b>" & strQuestionDesc & "</b></span><br><br></td>")
        '                sb.Append("              </tr>")


        '                For intIndex = LBound(arrQuestions) To UBound(arrQuestions)

        '                    If Len(arrQuestions(intIndex)) > 0 Then

        '                        If intTotalVotes > 0 Then

        '                            dblPercentage = arrAnswers(intIndex) / intTotalVotes * 100

        '                        End If

        '                        sb.Append("            <tr>")
        '                        sb.Append("              <td width=""110""><div align=""left""><span class=""Estilo1"">" & arrQuestions(intIndex) & "</span></div></td>")
        '                        sb.Append("              <td width=""40""  valign=bottom  align=right class=""Estilo1"">" & FormatNumber(dblPercentage, 2) & "%</td>")
        '                        sb.Append("            </tr>")

        '                        sb.Append("            <tr><td colspan=2 height=3></td></tr>")

        '                    End If

        '                Next


        '                'Show Total Number of Votes
        '                If intTotalVotes > 0 Then

        '                    sb.Append("            <tr>")
        '                    sb.Append("              <td colspan=2 class=""Estilo1"">N&#250;mero de votos: " & intTotalVotes & "</td>")
        '                    sb.Append("            </tr>")

        '                End If


        '                sb.Append("          </table>")

        '            End If

        '        End If


        '        Return sb.ToString()

        '    Catch ex As Exception

        '        ShowError(ex)

        '    Finally

        '        reader = Nothing
        '        sb = Nothing

        '    End Try


        'End Function

        Public Function TestCase(ByVal strVar As String) As String

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As New StringBuilder("")
            Dim intID As Integer = 0
            Dim strFld1 As String = ""
            Dim strFld2 As String = ""
            Dim strFld3 As String = ""
            Dim strFld4 As String = ""
            Dim strFld5 As String = ""
            Dim strSQL As String = ""


            reader = ParaLideres.GenericDataHandler.GetRecords("Select id, first_name, last_name, city, state, main_language from reg_users order by id")

            Do While reader.Read()

                Try
                    strSQL = ""

                    intID = reader(0)
                    strFld1 = Functions.FormatString(reader(1))
                    strFld2 = Functions.FormatString(reader(2))
                    strFld3 = Functions.FormatString(reader(3))
                    strFld4 = Functions.FormatString(reader(4))
                    strFld5 = Functions.FormatString(reader(5))

                    strSQL = "update reg_users set first_name='" & Functions.FormatCase(strFld1, True) & "',last_name='" & Functions.FormatCase(strFld2, True) & "',city='" & Functions.FormatCase(strFld3, True) & "',state='" & Functions.FormatCase(strFld4, True) & "',main_language='" & Functions.FormatCase(strFld5, True) & "' Where id=" & intID

                    ParaLideres.GenericDataHandler.ExecSQL(strSQL)

                    sb.Append("<p>")
                    sb.Append(strSQL)

                Catch ex As Exception

                    sb.Append(ex.ToString())

                    Throw ex

                End Try



            Loop

            reader.Close()

            reader = Nothing

            Return sb.ToString

            sb = Nothing

        End Function

        Public Sub TestReg(Optional ByVal intID As Integer = 1)

            Dim objReg As Registration = New Registration

            OnLoadScript = "window_onload();"

            PageTitle = "Tu Perfil"
            PageContent = objReg.RegForm(intID)

            objReg = Nothing

        End Sub

        Public Sub DisplayForum(ByVal intForumId As Integer)

            Dim sb As New StringBuilder("")
            Dim objForum As Forum
            Dim intIndex As Integer = 1


            If Request("index") <> "" Then intIndex = CInt(Request("index"))

            objForum = New Forum(intForumId)

            sb.Append(objForum.ForumDesc)
            sb.Append("<p>" & objForum.DisplayForum(intIndex))

            'IsSimplePage = True
            PagePath = objForum.Trail
            PageTitle = objForum.ForumName

            PageContent = sb.ToString()

            objForum = Nothing

        End Sub

        Public Sub DisplayMessageForm(ByVal intForumId As Integer, Optional ByVal intParentId As Integer = 0)

            Dim sb As New StringBuilder("")
            Dim objForum As Forum

            Try

                objForum = New Forum(intForumId)

                sb.Append(objForum.DisplayReplyForm(intParentId))

                If intParentId = 0 Then

                    PageTitle = "Añadir Nuevo Mensaje"

                Else

                    PageTitle = "Responder Mensaje"

                End If

                PagePath = objForum.Trail
                PageContent = sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing
                objForum = Nothing

            End Try


        End Sub

        Public Sub DisplayThread(ByVal intMessageId As Integer, ByVal intForumId As Integer)

            If intMessageId = 0 Then Response.Redirect(_project_path & "forum.aspx?forum_id=" & intForumId)

            Dim objForum As Forum

            objForum = New Forum

            PageTitle = objForum.ForumName
            PageContent = objForum.DisplayThread(intMessageId)
            PagePath = objForum.Trail 'NOTE do not move this line of code This has to be executed after objForum.displaythread

            objForum = Nothing

        End Sub

        Public Sub PostMessage(ByVal intMessageId As Integer, ByVal intForumId As Integer, ByVal intReplyTo As Integer, ByVal strSubject As String, ByVal strBody As String)

            Dim objForum As Forum

            Dim strUserName As String = Functions.FormatCase(_objUser.getFirstName, True) & " " & Functions.FormatCase(_objUser.getLastName, True)

            Try

                objForum = New Forum

                objForum.PostMessage(intMessageId, intForumId, intReplyTo, strUserName, LCase(_objUser.getEmail), Functions.FormatCase(strSubject), strBody, Date.Today)

            Catch ex As Exception

                ShowError(ex)

            Finally

                objForum = Nothing

            End Try

        End Sub

        Public Function DropDownMenu() As String

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim readerSub As System.Data.SqlClient.SqlDataReader
            Dim sb As New StringBuilder("")

            Dim intSectionID As Integer = 0
            Dim strSectionName As String = ""

            reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMainSections")
            sb.Append("      <tr>")
            sb.Append("        <td height=""28"" valign=""top"" class=""Estilo1"">")

            'sb.Append("<img src=""" & _project_path & "_images/desplegable_image.jpg"" width=""146"" height=""15"">")

            sb.Append("      <form action=" & _project_path & "section.aspx method='POST' name=frmMenu id=frmMenu >" & Chr(13))

            sb.Append("      <table border=0 cellpaddin=0 cellspacing=0><tr><td>" & Chr(13))

            sb.Append("          <select name=""section_id""  onChange='document.frmMenu.submit();' class=""frmSelect"" >" & Chr(13))

            sb.Append("            <option>selecciona una opci&oacute;n</option>" & Chr(13))

            Trace.Warn("inside the form")

            If reader.HasRows Then

                Do While reader.Read()

                    intSectionID = reader(0)
                    strSectionName = reader(1)

                    sb.Append("<option value=" & intSectionID & " class=frmSelectBold>" & strSectionName & "</option>" & Chr(13))

                    readerSub = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionFirstTier " & intSectionID)

                    If readerSub.HasRows Then

                        Do While readerSub.Read()

                            intSectionID = readerSub(0)
                            strSectionName = readerSub(1)

                            sb.Append("<option value=" & intSectionID & ">--" & strSectionName & "</option>" & Chr(13))

                        Loop

                    End If

                Loop

            End If

            sb.Append("          </select><br>" & Chr(13))
            sb.Append("      </td></td></table>" & Chr(13))
            sb.Append("      </form>" & Chr(13))
            sb.Append("       </td>")
            sb.Append("      </tr>")

            reader.Close()
            readerSub.Close()

            Return sb.ToString()

        End Function

        Function Rollovers(ByVal arrItems() As String) As String

            Dim sb As New StringBuilder("")

            Dim strFileName As String = ""
            Dim strImageName As String = ""
            Dim strImageNameOver As String = ""
            Dim strJavaName As String = ""
            Dim strJavaNameOver As String = ""

            Dim i As Integer = 0

            'JavaScript Rollovers
            sb.Append("<Script LANGUAGE=""JavaScript"">" & Chr(13))
            sb.Append("<!--" & Chr(13))
            sb.Append("if (document.images) {" & Chr(13))

            For i = LBound(arrItems) To UBound(arrItems)

                strFileName = LCase(arrItems(i))

                strImageName = _project_path & "_images/menu_" & strFileName & ".gif"
                strImageNameOver = _project_path & "_images/menu_" & strFileName & "_o.gif"
                strJavaName = "menu_" & strFileName
                strJavaNameOver = "menu_" & strFileName & "_o"

                sb.Append("		" & strJavaName & " = new Image();" & Chr(13))
                sb.Append("		" & strJavaName & ".src = """ & strImageName & """;" & Chr(13))

                sb.Append("		" & strJavaNameOver & " = new Image();" & Chr(13))
                sb.Append("		" & strJavaNameOver & ".src = """ & strImageNameOver & """;" & Chr(13))
            Next

            sb.Append("}" & Chr(13))

            sb.Append("function imageRollover(imageID,imageSource) {" & Chr(13))
            sb.Append("if (document.images) {" & Chr(13))
            sb.Append("document.images[imageID].src = eval(imageSource + "".src"");" & Chr(13))

            sb.Append("}" & Chr(13))
            sb.Append("}" & Chr(13))

            'openwindow function
            sb.Append("function Openwindow(mURL)" & Chr(13))
            sb.Append("{" & Chr(13))
            sb.Append("window.open(mURL,"""",""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=640,height=480"");" & Chr(13))
            sb.Append("}" & Chr(13))


            sb.Append("// -->" & Chr(13))
            sb.Append("</script>" & Chr(13))

            Return sb.ToString()

        End Function

        Public Function GenerateMenuItem(ByVal strFileName As String, ByVal strLink As String) As String

            Dim sb As New StringBuilder("")

            Dim strImageName As String = ""
            Dim strImageNameOver As String = ""
            Dim strJavaName As String = ""
            Dim strJavaNameOver As String = ""

            strImageName = _project_path & "_images/menu_" & strFileName & ".gif"
            strImageNameOver = _project_path & "_images/menu_" & strFileName & "_o.gif"
            strJavaName = "menu_" & strFileName
            strJavaNameOver = "menu_" & strFileName & "_o"

            sb.Append("<a href=""" & _project_path & strLink & """")
            sb.Append(" onmouseover=""imageRollover('" & strFileName & "', '" & strJavaNameOver & "'); return true;""")
            sb.Append("  onmouseout=""imageRollover('" & strFileName & "', '" & strJavaName & "'); return true;"">")
            sb.Append("<img src=" & strImageName & " name=" & strFileName & " border=0>")
            sb.Append("</a>")

            Return sb.ToString()

        End Function

        'Public Function GenerateMenu() As String

        '    Dim sb As New StringBuilder("")
        '    Dim intIndex As Integer = 0

        '    sb.Append("<tr>" & Chr(13))
        '    sb.Append("<td align=center>" & Chr(13))
        '    sb.Append(" <table width=""140"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13))

        '    'Header
        '    sb.Append("  <tr>" & Chr(13))
        '    sb.Append("   <td height=""1"" background=""" & _project_path & "_images/puntos_horizontales.gif""><img src=""" & _project_path & "_images/puntos_horizontales.gif"" width=""3"" height=""1"" align=""left""></td>" & Chr(13))
        '    sb.Append("  </tr>" & Chr(13))


        '    For intIndex = 0 To 5

        '        sb.Append("<tr>" & Chr(13))
        '        sb.Append("<td align=center>" & GenerateMenuItem(_arrItems(intIndex), _arrLinks(intIndex)) & "</td>" & Chr(13))
        '        sb.Append("</tr>" & Chr(13))

        '    Next

        '    sb.Append(" </table>" & Chr(13))
        '    sb.Append("</td>" & Chr(13))
        '    sb.Append("</tr>" & Chr(13))

        '    Return sb.ToString()

        'End Function

        Public Function GetTotalNoOfArticles(ByVal intSectionId As Integer) As Integer

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

                        Trace.Write("This section id: " & intThisSection & " total art: " & intThisTotal)

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

        Public Function LogonForm() As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim objFrm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("frmLogon")

            Dim strRefer As String = ""
            Dim strFldUsername As String = ""
            Dim strFldPassword As String = ""
            Dim strEmail As String = ""
            Dim strPassWord As String = ""

            Dim intRandom As Integer = ParaLideres.Functions.GetRandomNumber(500, 999)

            'This is to prevent dictionary and brute force attaks
            strFldUsername = "username" & intRandom
            strFldPassword = "password" & intRandom

            Try

                Try
                    strRefer = Request.UrlReferrer.PathAndQuery
                Catch ex As Exception
                End Try

                Try
                    strEmail = Request("email")
                Catch ex As Exception
                End Try

                If Request("url_redirect") <> "" Then

                    strRefer = Request("url_redirect")

                End If

                Trace.Write("Referrer: " & strRefer)

                If strEmail <> "" Then

                    Try
                        strPassWord = ParaLideres.GenericDataHandler.ExecScalar("proc_GetPasswordByEmail '" & strEmail & "'")
                    Catch ex As Exception
                    End Try

                    If strPassWord = "" Then

                        sb.Append("<font color=red>" & strEmail & " no es una direcci&#243;n que se ha registrado en Para L&#237;deres</font>.<br>")

                        sb.Append("Por favor reg&#237;strate para poder ver esta secci&#243;n de ParaLderes haciendo <a href=registration.aspx>clic aqu&#237;</a>.<p><b>Si ya te has registrado</b> entonces ingresa tu informaci&#243;n en la forma de abajo y presiona el bot&#243;n ""Enviar Informaci&#243;n"".")

                        strEmail = ""

                    Else

                        sb.Append("<font color=red>La clave que proporcionaste no es la correcta para esta cuenta</font>. <a href=emailpassword.aspx?email=" & strEmail & ">¿Has olvidado tu clave?</a>")

                    End If

                End If

                'DO NOT MOVE THIS UP
                If Not IsNothing(Request.Cookies("username")) And strEmail = "" Then

                    strEmail = Request.Cookies("username").Value

                End If


                sb.Append(objFrm.FormAction("chcklog.aspx"))
                sb.Append(objFrm.FormHidden("url_redirect", strRefer))
                sb.Append(objFrm.FormTextBox("E-mail", strFldUsername, strEmail, 100, "Ingresa el e-mail que usaste para registrarte", True, 9, False, True))
                sb.Append(objFrm.FormPasswordTextBox("Clave", strFldPassword, "", 16, "Ingresa tu clave secreta (4 a 16 characteres).", True, 4, False))
                sb.Append(objFrm.FormEnd("Enviar Informaci&#243;n"))

                sb.Append("<p align=center><a href=emailpassword.aspx?email=" & strEmail & ">¿Has olvidado tu clave?</a></p>" & Chr(13))

                sb.Append("<p align=center><a href=registration.aspx>Reg&#237;strate como nuevo usuario</a></p>" & Chr(13))

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing
                objFrm = Nothing

            End Try

        End Function

        Public Sub DisplaySectionPage(ByVal intSectionID As Integer)

            Dim objSection As ParaLideres.Section = New ParaLideres.Section(intSectionID)
            Dim intIndex As Integer = 1

            Try

                intIndex = CInt(Request("index"))

            Catch ex As Exception

            End Try

            Try

                If intIndex = 0 Then intIndex = 1

                'NOTE: do not change the order of this.  pagecontent must go first than page title to get totalarticles
                PagePath = objSection.GetPath()
                PageTitle = objSection.SectionName & " (" & objSection.TotalArticles & ")"
                PageContent = objSection.DisplaySection(intIndex)

            Catch ex As Exception

                ShowError(ex)

            Finally

                objSection = Nothing

            End Try

        End Sub

        Public Sub CreateSectionPageHtml(ByVal intSectionID As Integer, ByVal strDirPath As String)

            Dim objSection As ParaLideres.Section = New ParaLideres.Section(intSectionID)
            Dim file As System.IO.StreamWriter

            Dim intIndex As Integer = 1

            Try

                If Request("index") <> "" Then intIndex = CInt(Request("index"))

                file = New System.IO.StreamWriter(strDirPath & "\section" & intSectionID & ".htm", False, System.Text.Encoding.Unicode)

                'NOTE: do not change the order of this.  pagecontent must go first than page title to get totalarticles
                PagePath = objSection.GetPath()
                PageContent = objSection.DisplaySectionHtml(intIndex)
                PageTitle = objSection.SectionName & "(" & objSection.TotalArticles & ")"

                file.Write(PageTemplateHtml(objSection.SectionName))

            Catch ex As Exception

                ShowError(ex)

            Finally

                file.Close()
                file = Nothing

                objSection = Nothing

            End Try

        End Sub

        Public Sub CreateHtmlPagesForRobots()

            Dim file As System.IO.StreamWriter
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intPageId As Integer = 0
            Dim strPageTitle As String = ""

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("SELECT page_title, page_id FROM pages WHERE (isposted = 1) AND page_id = 42 ORDER BY typeofarticle")

                If reader.HasRows() Then

                    Do While (reader.Read())

                        strPageTitle = reader(0)

                        intPageId = reader(1)

                        file = New System.IO.StreamWriter(Request.PhysicalApplicationPath & "html\page" & intPageId & ".htm", False, System.Text.Encoding.Unicode)

                        DisplayArticle(intPageId, , , , False) 'this sets the value for pagetitle and pagecontent

                        file.Write(PageTemplateHtml(strPageTitle))
                        file.Close()

                    Loop

                End If

            Catch ex As Exception

                ShowError(ex)

            Finally

                reader.Close()
                reader = Nothing

                file = Nothing

            End Try

        End Sub

        Public Sub DisplayArticle(ByVal intArticleID As Integer, Optional ByVal intIndex As Integer = 0, Optional ByVal strReferrer As String = "", Optional ByVal strSearchParam As String = "", Optional ByVal blIsForAspPage As Boolean = True)

            Dim objArticle As DataLayer.pages
            Dim objSection As ParaLideres.Section
            Dim sb As StringBuilder = New StringBuilder("")


            Dim strFileImage As String = ""
            Dim strFileType As String = ""

            Dim blShowAuthorPicture As Boolean = True
            Dim blIsForEmail As Boolean = False

            Try

                If Request("send_email") = "yes" Then

                    blIsForEmail = True

                End If

                Trace.Write("-->strReferrer: " & strReferrer)

                objArticle = New DataLayer.pages(intArticleID)

                If (intArticleID > 0 And objArticle.getIsposted = 1) Then

                    objSection = New ParaLideres.Section(objArticle.getSectionId)

                    If Request("format") = "print" Then

                        blShowAuthorPicture = False

                        'PageFormats.PrintFormat 

                    End If


                    If objArticle.getTypeofarticle < 3 Then

                        sb.Append(ParaLideres.Functions.DisplayAuthorInfo(objArticle.getUserId, objArticle.getPosted, blShowAuthorPicture))

                        'If Request("format") = "" Then sb.Append(ParaLideres.Functions.DisplayArticleRating(objArticle.getPageId, objArticle.getSectionId, objSection.SectionName, intIndex))

                        sb.Append("</br>")

                    End If

                    Select Case objArticle.getTypeofarticle()

                        Case 1, 0 'Normal

                            sb.Append("<p>")

                            If objArticle.getPic <> "" Then

                                If System.IO.File.Exists(Server.MapPath("files/" & objArticle.getPic)) Then

                                    Select Case System.IO.Path.GetExtension(objArticle.getPic)

                                        Case ".jpg", ".gif"

                                            sb.Append("<img src=" & _project_path & "files/" & objArticle.getPic & " align=left>")

                                    End Select

                                End If

                            End If

                            sb.Append(Functions.ReplaceCR(objArticle.getBody))

                            sb.Append("</p>")

                        Case 2 'Word Format of Pdf

                            sb.Append("<p>" & Functions.ReplaceCR(objArticle.getBlurb) & "</p>")

                            Trace.Write("IO.Path: " & System.IO.Path.GetExtension(objArticle.getPic))

                            Select Case System.IO.Path.GetExtension(objArticle.getPic)

                                Case ".doc"

                                    strFileImage = "<img src=" & ProjectPath & "images/word.gif border=0>"

                                    strFileType = "Baja art&#237;culo en MS Word&reg;"


                                Case ".pdf"

                                    strFileImage = "<img src=" & ProjectPath & "images/pdf.gif border=0>"

                                    strFileType = "Baja art&#237;culo en Adobe Acrobat&reg;"

                            End Select

                            If strFileImage <> "" Then sb.Append("<br><center><a href=" & ProjectPath & "files/" & objArticle.getPic & ">" & strFileImage & " " & strFileType & "</a></center>")

                        Case 3 'other pages

                            'Me.IsSimplePage = True

                            sb.Append(Functions.ReplaceCR(objArticle.getBody))

                    End Select

                    If objArticle.getTypeofarticle < 3 And ((Request("format") = "") Or (Request("close") = "y")) And blIsForAspPage Then

                        If InStr(strReferrer, "section.aspx") > 0 Then

                            sb.Append("<p align=center>| <a href=" & ProjectPath & "section.aspx?section_id=" & objArticle.getSectionId & "&index=" & intIndex & ">Regresar a " & objSection.SectionName & "</a> | </p>")

                        ElseIf InStr(strReferrer, "destacado.aspx") > 0 Then

                            sb.Append("<p align=center>| <a href=" & ProjectPath & "destacado.aspx?index=" & intIndex & ">Regresar a Destacado</a> | </p>")

                        ElseIf InStr(strReferrer, "lo_ultimo.aspx") > 0 Then

                            sb.Append("<p align=center>| <a href=" & ProjectPath & "lo_ultimo.aspx?index=" & intIndex & ">Regresar a Lo Ultimo</a> | </p>")

                        ElseIf InStr(strReferrer, "mis_favoritos.aspx") > 0 Then

                            sb.Append("<p align=center>| <a href=" & ProjectPath & "mis_favoritos.aspx?index=" & intIndex & ">Regresar a Mis Favoritos</a> | </p>")

                        ElseIf InStr(strReferrer, "search.aspx") > 0 And strSearchParam <> "" Then

                            sb.Append("<p align=center>| <a href=" & ProjectPath & "search.aspx?index=" & intIndex & "&shearchparam=" & strSearchParam & ">Regresar a Resultados de B&#250;squeda</a> | </p>")

                        End If

                        'Link to rate this article if not voted already
                        If (objArticle.getRating = 1) Then

                            sb.Append(Rate(intArticleID, blIsForEmail))

                        End If

                        sb.Append(Tags(objArticle.getPageId, objArticle.getKeywords, blIsForEmail))

                        PagePath = objSection.GetPath(intIndex)

                    End If

                    If blIsForAspPage Then

                        If IsNothing(Cache.Get("LastViewed")) Then


                            'sb.Append("<a href=""view.aspx?" & strLink & """ onmouseover=""ShowAjaxContent('view_pic.aspx?" & strLink & "',500,350,this);"" ")
                            'sb.Append(" onmouseout=""ClearError(divAjaxContent);""")
                            'sb.Append(">")

                            '_arrLastViewed(0, 0) = "<br><img src=""_images/item.gif"" valign=""middle""><a href=""" & ProjectPath & "article.aspx?page_id=" & intArticleID & """  onclick=""ShowAjaxContent('article_preview.aspx?page_id=" & intArticleID & "&format=print&close=y',530,200,this);"" >" & objArticle.getPageTitle & "</a>"

                            _arrLastViewed(0, 0) = "<br><img src=""_images/item.gif"" valign=""middle""><a href=""javascript:""  onclick=""ShowAjaxContent('article_preview.aspx?page_id=" & intArticleID & "&format=print&close=y',530,200,this);"" >" & objArticle.getPageTitle & "</a>"

                            'onmouseout=""ClearError(divAjaxContent);""

                            _arrLastViewed(0, 1) = Request.UserHostAddress() 'Request.ServerVariables("RemoteAddresss")

                            _arrLastViewed(0, 2) = Date.Today()

                            _arrLastViewed(0, 3) = intArticleID

                            Cache.Insert("LastViewed", _arrLastViewed)

                            Cache.Insert("LastViewedIndex", 1)

                        Else 'If IsNothing(Cache.Get("LastViewed")) Then

                            Dim intI As Integer = 0
                            Dim intX As Integer = 0
                            Dim blFound As Boolean = False

                            _arrLastViewed = Cache.Get("LastViewed")
                            intI = Cache.Get("LastViewedIndex")

                            For intX = 0 To _arrLastViewed.GetUpperBound(0)

                                If intArticleID = CInt(_arrLastViewed(intX, 3)) Then blFound = True

                            Next

                            If Not blFound Then

                                '_arrLastViewed(intI, 0) = "<br><img src=""_images/item.gif"" valign=""middle""><a href=""" & ProjectPath & "article.aspx?page_id=" & intArticleID & """  onclick=""ShowAjaxContent('article_preview.aspx?page_id=" & intArticleID & "&format=print&close=y',530,200,this);""  >" & objArticle.getPageTitle & "</a>"

                                _arrLastViewed(intI, 0) = "<br><img src=""_images/item.gif"" valign=""middle""><a href=""javascript:""  onclick=""ShowAjaxContent('article_preview.aspx?page_id=" & intArticleID & "&format=print&close=y',530,200,this);""  >" & objArticle.getPageTitle & "</a>"

                                'onmouseout=""ClearError(divAjaxContent);"" 

                                _arrLastViewed(intI, 1) = Request.UserHostAddress() 'Request.ServerVariables("REMOTE_ADDR")

                                _arrLastViewed(intI, 2) = Date.Today()

                                _arrLastViewed(intI, 3) = intArticleID

                                intI = intI + 1

                                If intI > 9 Then intI = 0

                                Cache.Insert("LastViewed", _arrLastViewed)

                                Cache.Insert("LastViewedIndex", intI)

                            End If

                        End If 'If IsNothing(Cache.Get("LastViewed")) Then

                    End If 'If blIsForAspPage Then

                Else 'If intArticleID > 0 And objArticle.getIsposted Then

                    sb.Append("El art&#237;culo que buscas no ha sido publicado todav&#237;a.  Regresa pronto para poder leerlo.")

                End If 'If intArticleID > 0 And objArticle.getIsposted Then


                PageTitle = objArticle.getPageTitle
                PageContent = sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                If Not IsNothing(objSection) Then objSection = Nothing
                objArticle = Nothing
                sb = Nothing

            End Try

        End Sub


        Public Sub DisplayArticleV4(ByVal intArticleID As Integer, Optional ByVal intIndex As Integer = 0, Optional ByVal strReferrer As String = "", Optional ByVal strSearchParam As String = "", Optional ByVal blIsForAspPage As Boolean = True)

            Dim objArticle As DataLayer.pages
            Dim objSection As ParaLideres.Section
            Dim objAuthor As DataLayer.reg_users
            Dim sb As StringBuilder = New StringBuilder("")


            Dim strAuthorName As String = ""

            Dim strFileImage As String = ""
            Dim strFileType As String = ""

            Dim blIsForEmail As Boolean = False

            Try

                If Request("send_email") = "yes" Then

                    blIsForEmail = True

                End If

                Trace.Write("-->strReferrer: " & strReferrer)

                objArticle = New DataLayer.pages(intArticleID)

                If (intArticleID > 0 And objArticle.getIsposted = 1) Then

                    objAuthor = New DataLayer.reg_users(objArticle.getUserId)

                    objSection = New ParaLideres.Section(objArticle.getSectionId)

                    strAuthorName = objAuthor.getFirstName & " " & objAuthor.getLastName

                    'sb.Append("        	 <!-- start bread crumbs -->" & vbCrLf)

                    'sb.Append("            <div id=""breadcrumb"" class=""clearfix breadcrumbs"">" & vbCrLf)

                    'sb.Append(objSection.GetPath(intIndex))

                    'sb.Append("            </div>" & vbCrLf)

                    'sb.Append("            <!-- ends bread crumbs -->" & vbCrLf)

                    'sb.Append("            <!-- Central Content box -->" & vbCrLf)

                    'sb.Append("            <div id=""central_content"" class=""clearfix"">" & vbCrLf)

                    'sb.Append("                <!-- Start Resource box -->" & vbCrLf)

                    sb.Append("                <div id=""resource_box"" class=""clearfix"">" & vbCrLf)

                    sb.Append("                	<div id=""r_titlebar"" class=""left"">" & vbCrLf)

                    sb.Append(Functions.ShowPicture(objArticle.getUserId, objAuthor.getPicture, objAuthor.getSex, "r_author_pic"))

                    sb.Append("                        <div>" & vbCrLf)

                    sb.Append("                        <h1>" & objArticle.getPageTitle & "</h1>" & vbCrLf)

                    sb.Append("                            <span class=""r_author_name"">Por: <a href=""" & _project_path & "bio.aspx?user_id=" & objAuthor.getId & """ class=""author_link"" title=""" & strAuthorName & """>" & Functions.camelNotation(strAuthorName, False) & "</a></span>" & vbCrLf)

                    sb.Append("                            <!-- starts rating -->" & vbCrLf)

                    sb.Append("                            <div class=""rating r_rating"">" & vbCrLf)

                    'Link to rate this article if not voted already
                    If (objArticle.getRating = 1) Then

                        sb.Append(Rate(intArticleID, blIsForEmail))

                    End If

                    sb.Append("                            </div>" & vbCrLf)

                    sb.Append("                            <!-- ends rating-->" & vbCrLf)

                    sb.Append("                        </div>" & vbCrLf)

                    sb.Append("                    </div>" & vbCrLf)

                    sb.Append("                    <div id=""r_sharebar"" class=""left"">" & vbCrLf)

                    sb.Append("                    	<div class=""left"">" & vbCrLf)

                    If objArticle.getTypeofarticle = 2 Then

                        Trace.Write("IO.Path: " & System.IO.Path.GetExtension(objArticle.getPic))

                        Select Case System.IO.Path.GetExtension(objArticle.getPic)

                            Case ".doc"

                                strFileImage = "<img src=""" & ProjectPath & "images/word.gif"" alt=""Descargar Recurso"" title=""Descargar Recurso"">"

                                strFileType = "Baja art&#237;culo en MS Word&reg;"


                            Case ".pdf"

                                strFileImage = "<img src=""" & ProjectPath & "images/pdf.gif"" alt=""Descargar Recurso"" title=""Descargar Recurso"">"

                                strFileType = "Baja art&#237;culo en Adobe Acrobat&reg;"

                        End Select

                        If strFileImage <> "" And objArticle.getPic <> "" Then

                            sb.Append("                            <a href=""" & ProjectPath & "files/" & objArticle.getPic & """ class=""r_download left"">" & vbCrLf)

                            sb.Append(strFileImage)

                            sb.Append("                            </a>" & vbCrLf)

                        End If

                    End If

                    sb.Append("                        </div>" & vbCrLf)

                    sb.Append("                        <div class=""right"">" & vbCrLf)

                    'sb.Append("                        	<script type=""text/javascript""> var addthis_config = { username: ""paralideres""; ui_language: ""es"" } </script>" & vbCrLf)

                    sb.Append("                            <div class=""addthis_toolbox addthis_default_style"">" & vbCrLf)

                    'sb.Append("                                <a class=""addthis_button_twitter"">Retweet</a>" & vbCrLf)

                    'sb.Append("                                <a class=""addthis_button_facebook"">Compartir</a>" & vbCrLf)

                    'sb.Append("                                <a class=""addthis_button_email"">Email</a>" & vbCrLf)

                    'sb.Append("                                <a class=""addthis_button_compact"">más ...</a>" & vbCrLf)

                    sb.Append(CompartirArticulo(intArticleID, blIsForEmail))

                    sb.Append("                            </div>" & vbCrLf)

                    sb.Append("                        </div>" & vbCrLf)

                    sb.Append("                    </div>" & vbCrLf)

                    sb.Append("                 	<div id=""r_content"">" & vbCrLf)

                    sb.Append("                        <p class=""r_content_txt"">" & vbCrLf)

                    If objArticle.getTypeofarticle = 1 Then

                        If objArticle.getPic <> "" Then

                            If System.IO.File.Exists(Server.MapPath("files/" & objArticle.getPic)) Then

                                Select Case System.IO.Path.GetExtension(objArticle.getPic)

                                    Case ".jpg", ".gif"

                                        sb.Append("<img src=""" & _project_path & "files/" & objArticle.getPic & """ align=""left"" />")

                                End Select

                            End If

                        End If

                    End If

                    sb.Append("<br />" & Functions.ReplaceCR(objArticle.getBlurb))

                    sb.Append("<br />" & Functions.ReplaceCR(objArticle.getBody))

                    sb.Append("                        </p>" & vbCrLf)

                    sb.Append("                        <p class=""r_tags"">")

                    sb.Append(Etiquetas(objArticle.getKeywords))


                    'LINKS FOR NAVIGATION
                    If objArticle.getTypeofarticle < 3 And ((Request("format") = "") Or (Request("close") = "y")) And blIsForAspPage Then

                        If InStr(strReferrer, "section.aspx") > 0 Then

                            sb.Append("<br /><a href=" & ProjectPath & "section.aspx?section_id=" & objArticle.getSectionId & "&index=" & intIndex & ">Regresar a " & objSection.SectionName & "</a>")

                        ElseIf InStr(strReferrer, "destacado.aspx") > 0 Then

                            sb.Append("<br /><a href=" & ProjectPath & "destacado.aspx?index=" & intIndex & ">Regresar a Destacado</a>")

                        ElseIf InStr(strReferrer, "lo_ultimo.aspx") > 0 Then

                            sb.Append("<br /><a href=" & ProjectPath & "lo_ultimo.aspx?index=" & intIndex & ">Regresar a Lo Ultimo</a>")

                        ElseIf InStr(strReferrer, "mis_favoritos.aspx") > 0 Then

                            sb.Append("<br /><a href=" & ProjectPath & "mis_favoritos.aspx?index=" & intIndex & ">Regresar a Mis Favoritos</a>")

                        ElseIf InStr(strReferrer, "search.aspx") > 0 And strSearchParam <> "" Then

                            sb.Append("<br /><a href=" & ProjectPath & "search.aspx?index=" & intIndex & "&shearchparam=" & strSearchParam & ">Regresar a Resultados de B&#250;squeda</a>")

                        End If

                    End If

                    sb.Append("                        </p>" & vbCrLf)

                    sb.Append("                    </div>" & vbCrLf)


                    '-----------------------------------------------------------------------------------------------------------------------------------------------
                    'MORE ARTICLES FOR THIS USER

                    Dim reader As System.Data.SqlClient.SqlDataReader

                    Dim intPageId As Integer = 0

                    Dim strTitle As String = ""

                    Dim dtPosted As Date = #1/1/0900#


                    Try

                        'this gets the best top 4 rates articles for this author excluding the current page/article
                        reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetTop4PagesByUser " & objAuthor.getId & "," & intArticleID)

                        If reader.HasRows Then


                            sb.Append("                    <div id=""r_moreincat"">" & vbCrLf)

                            sb.Append("                        <div id=""tabs"">" & vbCrLf)

                            sb.Append("                            <ul class=""tabs-nav"">" & vbCrLf)

                            sb.Append("                                <li class=""tabs-title""><span>Más artículos por " & strAuthorName & "</span></li>" & vbCrLf)

                            sb.Append("                            </ul>" & vbCrLf)

                            sb.Append("                            <div>" & vbCrLf)


                            Do While reader.Read()


                                intPageId = reader(0)

                                strTitle = reader(1)

                                dtPosted = reader(2)


                                sb.Append("                                <div class=""lu_content_box"">" & vbCrLf)

                                sb.Append("                                    <img class=""lu_doc_icon"" src=""" & ProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

                                sb.Append("                                    <h1><a href=""" & ProjectPath & "article.aspx?page_id=" & intPageId & """>" & strTitle & "</a></h1>" & vbCrLf)

                                sb.Append("                                    <span>" & Functions.FormatHispanicDateTime(dtPosted) & "</span>" & vbCrLf)

                                sb.Append("                                </div>" & vbCrLf)

                            Loop


                            sb.Append("                            </div>" & vbCrLf)

                            sb.Append("                        </div>" & vbCrLf)

                            sb.Append("                    <div class=""tabs_more"">" & vbCrLf)

                            sb.Append("                            <a href=""" & ProjectPath & "pages_by_user.aspx?user_id=" & objAuthor.getId & """>Ver Más</a>" & vbCrLf)

                            sb.Append("                        </div>" & vbCrLf)

                            sb.Append("                	</div>" & vbCrLf)


                        End If


                    Catch ex As Exception

                        Throw ex

                    Finally

                        reader.Close()

                        reader = Nothing

                    End Try


                    sb.Append("                </div>" & vbCrLf)

                    sb.Append("                <!-- end Resource box -->" & vbCrLf)

                    'sb.Append("            </div>" & vbCrLf) 'end div center_content


                Else 'If intArticleID > 0 And objArticle.getIsposted Then

                    sb.Append("El art&#237;culo que buscas no ha sido publicado todav&#237;a.  Regresa pronto para poder leerlo.")

                End If 'If intArticleID > 0 And objArticle.getIsposted Then

                PagePath = objSection.GetPath(intIndex)
                PageContent = sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                If Not IsNothing(objSection) Then objSection = Nothing

                If Not IsNothing(objAuthor) Then objAuthor = Nothing

                If Not IsNothing(objArticle) Then objArticle = Nothing

                sb = Nothing

            End Try

        End Sub

        Public Sub DisplayArticleSansOptions(ByVal intArticleID As Integer)

            Dim objArticle As DataLayer.pages = New DataLayer.pages(intArticleID)
            Dim objSection As ParaLideres.Section = New ParaLideres.Section(objArticle.getSectionId)
            Dim sb As StringBuilder = New StringBuilder("")

            Dim strSessionName As String = ""
            Dim strFileImage As String = ""
            Dim strFileType As String = ""
            Dim blShowAuthorPicture As Boolean = True

            Try

                sb.Append("<p>")

                If objArticle.getPic <> "" Then

                    Dim flImage As System.IO.File

                    If flImage.Exists(Server.MapPath("/uploads/" & objArticle.getPic)) Then

                        Select Case System.IO.Path.GetExtension(objArticle.getPic)

                            Case ".jpg", ".gif"

                                sb.Append("<img src=/uploads/" & objArticle.getPic & " align=left>")

                        End Select



                    End If

                End If

                sb.Append(Functions.ReplaceCR(objArticle.getBody))

                sb.Append("</p>")

                'If objArticle.getTypeofarticle < 3 Then PagePath = objSection.GetPath()

                PageTitle = objArticle.getPageTitle
                PageContent = sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                objSection = Nothing
                objArticle = Nothing
                sb = Nothing

            End Try

        End Sub

        Public Function PublishMaterialsForm() As String

            If _objUser.getId > 0 Then

                Dim sb As StringBuilder = New StringBuilder("")
                Dim objFrm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("frmPublish")

                Try

                    sb.Append("Ingresa los datos de tu colaboraci&#243;n y luego presiona el bot&#243;n ""Enviar Colaboraci&#243;n"".")

                    sb.Append("<p>Tu colaboraci&#243;n no va a ser publicada inmediatamente ya que primero debe ser aprobada por el equipo de Para L&#237;deres. Este proceso tomar&#225; uno o dos d&#237;as.")

                    sb.Append(objFrm.FormAction("post_publish.aspx", , True))

                    sb.Append(objFrm.FormTextBox("T&#237;tulo", "Title", "", 150, "Ingresa el t&#237;tulo de tu colaboraci&#243;n", True))

                    sb.Append(objFrm.FormTextArea("Resumen", "Blurb", "", 5, 80, "Ingresa un pequeño resumen de tu colaboraci&#243;n (m&#225;ximo 255 caracteres.)", True, 500))

                    sb.Append(objFrm.FormTextArea("Texto", "Body", "", 20, 80, "Ingresa aqu&#237; el texto de tu colaboraci&#243;n.", False, 15000))

                    sb.Append(objFrm.FormSelect("Publicar En", "Section", "", "proc_GetSectionsForColaboracionesDropDown", "Selecciona la secci&#243;n donde este art&#237;culo aparecer&#225;.", False))

                    sb.Append(objFrm.FormFile("Subir Archivo", "File", 50, "Si tienes el documento en formato MS Word ® (.doc, .docx) o Adobe Acrobat® (.pdf) o Presentacion en Powerpoint(.ppt,.pptx) lo puedes colocar en nuestro sitio aqu&#237;.", False, ".doc,.docx,.pdf,.ppt,.pptx"))

                    sb.Append(objFrm.FormEnd("Enviar Colaboraci&#243;n"))

                    Return sb.ToString()

                Catch ex As Exception

                    ShowError(ex)

                Finally

                    sb = Nothing
                    objFrm = Nothing

                End Try

            Else

                Response.Redirect(_project_path & "logon.aspx?url_redirect=publish_my_materials.aspx")

            End If

        End Function

        Public Function PublishBlogForm() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                If _objUser.getId > 0 And _objUser.getSecurityLevel > 1 Then


                    Dim objFrm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("frmPublish")

                    Try

                        'sb.Append("Ingresa los datos de tu colaboraci&#243;n y luego presiona el bot&#243;n ""Enviar Colaboraci&#243;n"".")

                        'sb.Append("<p>Tu colaboraci&#243;n no va a ser publicada inmediatamente ya que primero debe ser aprobada por el equipo de Para L&#237;deres. Este proceso tomar&#225; uno o dos d&#237;as.")

                        sb.Append(objFrm.FormAction("post_blog.aspx", , True))

                        sb.Append(objFrm.FormTextBox("T&#237;tulo", "Title", "", 150, "Ingresa el t&#237;tulo de tu blog", True))

                        sb.Append(objFrm.FormTextArea("Resumen", "Blurb", "", 5, 80, "Ingresa un pequeño resumen de tu blog. (m&#225;ximo 255 caracteres.)", True, 500))

                        sb.Append(objFrm.FormTextArea("Texto", "Body", "", 20, 80, "Ingresa aqu&#237; el texto de tu blog.", False, 15000))

                        'sb.Append(objFrm.FormSelect("Publicar En", "Section", "", "proc_GetSectionsForColaboracionesDropDown", "Selecciona la secci&#243;n donde este art&#237;culo aparecer&#225;.", False))

                        
                        sb.Append(objFrm.FormEnd("Enviar Blog"))

                    Catch ex As Exception

                        Trace.Warn(ex.ToString())

                        ShowError(ex)

                    Finally


                        objFrm = Nothing

                    End Try

                ElseIf _objUser.getSecurityLevel > 1 Then

                    Response.Redirect(_project_path & "logon.aspx?url_redirect=publish_blog.aspx")

                Else

                    sb.Append("Tu no tienes acceso a esta secci&#243;n de Para L&#237;deres.")

                End If

                Return sb.ToString()

            Catch ex As Exception

                ShowError(ex)

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function RateArticle(ByVal intArticleId As Integer, ByVal intIndex As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim arrValues As String() = Split("|Malo|Regular|Bueno|Muy Bueno|Excelente", "|")
            Dim intX As Integer = 0
            Dim strFormName As String = "frmRating"
            Dim strFldName As String = "stars"

            Try

                sb.Append("<form action=post_rating.aspx method=post name=" & strFormName & " id=" & strFormName & ">")

                sb.Append("<input type=hidden name=article_id value=""" & intArticleId & """>")
                sb.Append("<input type=hidden name=index value=""" & intIndex & """>")

                sb.Append("<table cellspacing=0 cellpadding=3 border=0 width=250>")  '<tr><td><br><br></td></tr><tr><td align=right valign=top><font face="Verdana, Arial, Helvetica, Sans-Serif" size=2><b>Evalua este articulo:</b></td><td><b><font face="Verdana, Arial, Helvetica, Sans-Serif" size=2><input type=radio name=stars value=1><img src=/images/stars1.gif border=0>Malo<br><input type=radio name=stars value=2><img src=/images/stars2.gif border=0>Regular<br><input type=radio name=stars value=3><img src=/images/stars3.gif border=0>Bueno<br><input type=radio name=stars value=4><img src=/images/stars4.gif border=0>Muy Bueno<br><input type=radio name=stars value=5><img src=/images/stars5.gif border=0>Excelente<br></b><input type=submit value="Enviar" name=b1></td></tr></table></form></td></tr></table>

                sb.Append("<tr>")
                sb.Append("<td colspan=2 align=center class=estilo1>Eval&#250;a Este Art&#237;culo</td>")
                sb.Append("</tr>")

                For intX = 1 To 5


                    sb.Append("<tr>")
                    sb.Append("<td nowrap align=""right"">")
                    sb.Append("<input type=radio name=" & strFldName & " value=" & intX & "><img src=" & _project_path & "images/stars" & intX & ".gif border=0>")
                    sb.Append("</td>")
                    sb.Append("<td><div align=""left""><span class=""Estilo1"">" & arrValues(intX) & "</span></div></td>")
                    sb.Append("</tr>")

                Next

                sb.Append("<tr>")
                sb.Append("<td colspan=2 align=center><input type=button name=btnEval id=btnEval value=""Eval&#250;a"" alt=""Eval&#250;a""  onclick='VerifyForm" & strFormName & "();' class=frmbutton  height=25></td>")
                sb.Append("</tr>")

                sb.Append("</table>")

                sb.Append("</form>")

                sb.Append("<script language=javascript>" & Chr(13))
                sb.Append("<!--" & Chr(13))

                'function VerifyForm()
                sb.Append(Chr(13) & "function VerifyForm" & strFormName & "(){" & Chr(13))

                sb.Append("if (!RadioCheck" & strFormName & "('" & strFldName & "')){" & Chr(13))
                sb.Append("	alert('Por favor selecciona una de las opciones');" & Chr(13))
                sb.Append("}" & Chr(13))

                sb.Append("else {" & Chr(13))
                sb.Append("document." & strFormName & ".btnEval.value = ""Espera..."";" & Chr(13))
                sb.Append("document." & strFormName & ".btnEval.disabled = true;" & Chr(13))
                sb.Append("document." & strFormName & ".submit();" & Chr(13))
                sb.Append("}" & Chr(13)) 'end of second else statement

                'end verify form function
                sb.Append("}" & Chr(13))


                'validate radio buttons
                sb.Append("function RadioCheck" & strFormName & "(ps_fld){" & Chr(13))
                sb.Append("var ischecked = false;" & Chr(13))
                sb.Append("var num_of_items = 1;" & Chr(13))

                sb.Append("num_of_items = eval('document." & strFormName & ".' + ps_fld + '.length');" & Chr(13))

                sb.Append("	if (isNaN(num_of_items)) {" & Chr(13))
                sb.Append("		num_of_items = 1;" & Chr(13))
                sb.Append("	}" & Chr(13))


                'sb.Append("alert('num of items:' + num_of_items);")

                'if num = 1
                sb.Append("	if (num_of_items == 1) {" & Chr(13))

                sb.Append("	if (eval('document." & strFormName & ".' + ps_fld + '.checked') == true){" & Chr(13))
                sb.Append("		ischecked = true;" & Chr(13))
                sb.Append("	}" & Chr(13))

                sb.Append("	}" & Chr(13))

                ' else
                sb.Append("else	{" & Chr(13))

                sb.Append("for (i=0; i < num_of_items; i++){" & Chr(13))

                sb.Append("	if (eval('document." & strFormName & ".' + ps_fld + '[i].checked') == true){" & Chr(13))
                sb.Append("		ischecked = true;" & Chr(13))
                sb.Append("	}" & Chr(13))

                sb.Append("}" & Chr(13))

                sb.Append("}" & Chr(13)) 'end else

                sb.Append("return ischecked;" & Chr(13))

                sb.Append("}" & Chr(13))
                'end RadioCheck

                sb.Append("//-->" & Chr(13))
                sb.Append("</script>" & Chr(13))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function Tags(ByVal intPageId As Integer, ByVal strKeywords As String, ByVal blIsForEmail As Boolean) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strLink As String = "http://www.paralideres.org/article.aspx?page_id=" & intPageId

            Dim strPrint As String = ProjectPath & "article.aspx?page_id=" & intPageId & "&format=print"

            Dim strEmail As String = QueryStringUrl("&send_email=yes#email")

            Dim intX As Integer = 0

            'article.aspx?page_id=" & intArticleID & "&format=print target=new
            Try

                If strKeywords <> "" Then

                    Dim arrTags As String() = Split(strKeywords, " ")

                    sb.Append("<p align=""center"">Etiquetas:<br>" & vbCrLf)

                    For intX = 0 To UBound(arrTags)

                        sb.Append("<a href=""" & ProjectPath & "pages_by_tag.aspx?tag=" & arrTags(intX) & """ target=""new"">" & arrTags(intX) & "</a> ")

                    Next

                    'TODO añadir mis favoritos y mas articulos por este autor

                    'sb.Append("<p align=center>| ")

                    'sb.Append("<a href=pages_by_user.aspx?user_id=" & objArticle.getUserId)

                    'If InStr(strReferrer, "pages_by_user.aspx") > 0 Then sb.Append("&index=" & intIndex)

                    'sb.Append(">M&#225;s art&#237;culos publicados por este autor</a> | ")

                    'Link to print
                    'sb.Append("<a href=article.aspx?page_id=" & intArticleID & "&format=print target=new>Imprimir P&#225;gina <img src=" & _project_path & "images/print.gif border=0 align=bottom></a> | ")

                    'Link to add to my favorites
                    'If _objUser.getId > 0 Then sb.Append("<a href=post_my_favorites.aspx?page_id=" & intArticleID & "&index=" & intIndex & ">Añadir A 'Mis Favoritos'</a> | ")

                    'sb.Append("</p>")

                    sb.Append("</p>" & vbCrLf)

                    sb.Append("		<p align=""center"">Compartir esta artículo:<br>" & vbCrLf)

                    If _objUser.getId > 0 And Not blIsForEmail Then

                        sb.Append("		<a target='_blank' href='" & strEmail & "'><img src=""" & ProjectPath & "images/tags/ico_email_link2.png"" border=""0"" id=""enviar"" name=""enviar"" onmouseover=""document.enviar.src='" & ProjectPath & "images/tags/ico_email_link.png'"" onmouseout=""document.enviar.src='" & ProjectPath & "images/tags/ico_email_link2.png'"" alt=""Envia a un amigo por e-mail""/></a>" & vbCrLf)

                    End If

                    sb.Append("		<a target='_blank' href='" & strPrint & "'><img src=""" & ProjectPath & "images/tags/ico_imprimir2.png"" border=""0"" id=""print"" name=""print"" onmouseover=""document.print.src='" & ProjectPath & "images/tags/ico_imprimir.png'"" onmouseout=""document.print.src='" & ProjectPath & "images/tags/ico_imprimir2.png'"" alt=""Imprime el art&#237;culo""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href='http://www.facebook.com/share.php?u=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_facebook2.png"" border=""0"" id=""facebook"" name=""facebook"" onmouseover=""document.facebook.src='" & ProjectPath & "images/tags/ico_facebook.gif'"" onmouseout=""document.facebook.src='" & ProjectPath & "images/tags/ico_facebook2.png'"" alt=""Envia a facebook""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href='http://del.icio.us/post?title=&url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_delicious.gif"" border=""0"" id=""delicious"" name=""delicous"" onmouseover=""document.delicious.src='" & ProjectPath & "images/tags/ico_delicious.gif'"" onmouseover=""MM_swapImage('delicious','','http://www.futbolecuador.com/imagenes/icons/ico_delicious.gif',1)"" onmouseout=""document.delicious.src='" & ProjectPath & "images/tags/ico_delicious.gif'"" alt=""Envia a delicious""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href='http://myweb2.search.yahoo.com/myresults/bookmarklet?u=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_yahoo2.png"" border=""0"" id=""yahoo"" name=""yahoo"" onmouseover=""document.yahoo.src='" & ProjectPath & "images/tags/ico_yahoo.gif'"" onmouseout=""document.yahoo.src='" & ProjectPath & "images/tags/ico_yahoo2.png'"" alt=""Envia a Yahoo""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href='http://meneame.net/submit.php?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_meneame2.png"" border=""0"" id=""meneame"" name=""meneame"" onmouseover=""document.meneame.src='" & ProjectPath & "images/tags/ico_meneame.gif'"" onmouseout=""document.meneame.src='" & ProjectPath & "images/tags/ico_meneame2.png'"" alt=""Envia a meneame""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href='http://www.digg.com/submit?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_digg2.png"" border=""0"" id=""digg"" name=""digg"" onmouseover=""document.digg.src='" & ProjectPath & "images/tags/ico_digg.gif'"" onmouseout=""document.digg.src='" & ProjectPath & "images/tags/ico_digg2.png'"" alt=""Envia a digg""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href='http://www.mister-wong.es/addurl/?bm_url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_mister_wong2.png"" border=""0"" id=""mwong"" name=""mwong""  onmouseover=""document.mwong.src='" & ProjectPath & "images/tags/ico_mister_wong.gif'"" onmouseout=""document.mwong.src='" & ProjectPath & "images/tags/ico_mister_wong2.png'"" alt=""Envia a mister-wong""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href='http://www.google.com/bookmarks/mark?op=edit&bkmk=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_google2.png"" border=""0"" id=""google"" name=""google"" onmouseover=""document.google.src='" & ProjectPath & "images/tags/ico_google.png'"" onmouseout=""document.google.src='" & ProjectPath & "images/tags/ico_google2.png'"" alt=""Envia a google""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href='http://twitthis.com/twit?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_twitter2.png"" border=""0"" id=""twitthis"" name=""twitthis"" onmouseover=""document.twitthis.src='" & ProjectPath & "images/tags/ico_twitter.png'"" onmouseout=""document.twitthis.src='" & ProjectPath & "images/tags/ico_twitter2.png'"" alt=""Envia a twitter""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href='http://technorati.com/faves?add=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_technorati2.png"" border=""0"" id=""technorati"" name=""technorati"" onmouseover=""document.technorati.src='" & ProjectPath & "images/tags/ico_technorati2.png'"" onmouseout=""document.technorati.src='" & ProjectPath & "images/tags/ico_technorati2.png'"" alt=""Envia a technorati""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href=""http://www.myspace.com/Modules/PostTo/Pages/?u=" & strLink & """ ><img src=""" & ProjectPath & "images/tags/3.jpg"" border=""0"" id=""myspace"" name=""myspace"" onmouseover=""document.myspace.src='" & ProjectPath & "images/tags/3.jpg'"" onmouseout=""document.myspace.src='" & ProjectPath & "images/tags/3.jpg'"" alt=""Envia a myspace""/></a>" & vbCrLf)

                    sb.Append("		<a target='_blank' href=""https://favorites.live.com/quickadd.aspx?marklet=1&amp;mkt=en-us&amp;url=" & strLink & """><img src=""" & ProjectPath & "images/tags/8.jpg"" border=""0"" id=""live"" name=""live"" onmouseover=""document.live.src='" & ProjectPath & "images/tags/8.jpg'"" onmouseout=""document.live.src='" & ProjectPath & "images/tags/8.jpg'"" alt=""Envia a live""/></a>" & vbCrLf)

                    'sb.Append("		<a target='_blank' href=""http://www.stumbleupon.com/submit?url=" & strLink & """><img src=""" & ProjectPath & "images/tags/9.jpg"" border=""0"" id=""stumbleupon"" name=""stumbleupon"" onmouseover=""document.stumbleupon.src='" & ProjectPath & "images/tags/9.jpg'"" onmouseout=""document.stumbleupon.src='" & ProjectPath & "images/tags/9.jpg'"" alt=""Envia a stumbleupon""/></a>" & vbCrLf)

                    sb.Append("		</p>" & vbCrLf)


                End If 'If strKeywords <> "" Then

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function Etiquetas(ByVal strKeywords As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim intX As Integer = 0

            Try

                If strKeywords <> "" Then

                    Dim arrTags As String() = Split(strKeywords, " ")

                    sb.Append("Etiquetas: " & vbCrLf)

                    For intX = 0 To UBound(arrTags)

                        sb.Append("<a href=""" & ProjectPath & "pages_by_tag.aspx?tag=" & arrTags(intX) & """ target=""new"">" & arrTags(intX) & "</a> ")

                    Next

                End If 'If strKeywords <> "" Then

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function CompartirArticulo(ByVal intPageId As Integer, ByVal blIsForEmail As Boolean) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strLink As String = "http://www.paralideres.org/article.aspx?page_id=" & intPageId

            Dim strPrint As String = ProjectPath & "article.aspx?page_id=" & intPageId & "&format=print"

            Dim strEmail As String = QueryStringUrl("&send_email=yes#email")

            Try

                If _objUser.getId > 0 And Not blIsForEmail Then

                    sb.Append("		<a target='_blank' href='" & strEmail & "'><img src=""" & ProjectPath & "images/tags/ico_email_link2.png"" border=""0"" id=""enviar"" name=""enviar"" onmouseover=""document.enviar.src='" & ProjectPath & "images/tags/ico_email_link.png'"" onmouseout=""document.enviar.src='" & ProjectPath & "images/tags/ico_email_link2.png'"" alt=""Envia a un amigo por e-mail""/></a>" & vbCrLf)

                End If

                sb.Append("		<a target='_blank' href='" & strPrint & "'><img src=""" & ProjectPath & "images/tags/ico_imprimir2.png"" border=""0"" id=""print"" name=""print"" onmouseover=""document.print.src='" & ProjectPath & "images/tags/ico_imprimir.png'"" onmouseout=""document.print.src='" & ProjectPath & "images/tags/ico_imprimir2.png'"" alt=""Imprime el art&#237;culo""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href='http://www.facebook.com/share.php?u=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_facebook2.png"" border=""0"" id=""facebook"" name=""facebook"" onmouseover=""document.facebook.src='" & ProjectPath & "images/tags/ico_facebook.gif'"" onmouseout=""document.facebook.src='" & ProjectPath & "images/tags/ico_facebook2.png'"" alt=""Envia a facebook""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href='http://del.icio.us/post?title=&url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_delicious.gif"" border=""0"" id=""delicious"" name=""delicous"" onmouseover=""document.delicious.src='" & ProjectPath & "images/tags/ico_delicious.gif'"" onmouseover=""MM_swapImage('delicious','','http://www.futbolecuador.com/imagenes/icons/ico_delicious.gif',1)"" onmouseout=""document.delicious.src='" & ProjectPath & "images/tags/ico_delicious.gif'"" alt=""Envia a delicious""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href='http://myweb2.search.yahoo.com/myresults/bookmarklet?u=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_yahoo2.png"" border=""0"" id=""yahoo"" name=""yahoo"" onmouseover=""document.yahoo.src='" & ProjectPath & "images/tags/ico_yahoo.gif'"" onmouseout=""document.yahoo.src='" & ProjectPath & "images/tags/ico_yahoo2.png'"" alt=""Envia a Yahoo""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href='http://meneame.net/submit.php?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_meneame2.png"" border=""0"" id=""meneame"" name=""meneame"" onmouseover=""document.meneame.src='" & ProjectPath & "images/tags/ico_meneame.gif'"" onmouseout=""document.meneame.src='" & ProjectPath & "images/tags/ico_meneame2.png'"" alt=""Envia a meneame""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href='http://www.digg.com/submit?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_digg2.png"" border=""0"" id=""digg"" name=""digg"" onmouseover=""document.digg.src='" & ProjectPath & "images/tags/ico_digg.gif'"" onmouseout=""document.digg.src='" & ProjectPath & "images/tags/ico_digg2.png'"" alt=""Envia a digg""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href='http://www.mister-wong.es/addurl/?bm_url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_mister_wong2.png"" border=""0"" id=""mwong"" name=""mwong""  onmouseover=""document.mwong.src='" & ProjectPath & "images/tags/ico_mister_wong.gif'"" onmouseout=""document.mwong.src='" & ProjectPath & "images/tags/ico_mister_wong2.png'"" alt=""Envia a mister-wong""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href='http://www.google.com/bookmarks/mark?op=edit&bkmk=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_google2.png"" border=""0"" id=""google"" name=""google"" onmouseover=""document.google.src='" & ProjectPath & "images/tags/ico_google.png'"" onmouseout=""document.google.src='" & ProjectPath & "images/tags/ico_google2.png'"" alt=""Envia a google""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href='http://twitthis.com/twit?url=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_twitter2.png"" border=""0"" id=""twitthis"" name=""twitthis"" onmouseover=""document.twitthis.src='" & ProjectPath & "images/tags/ico_twitter.png'"" onmouseout=""document.twitthis.src='" & ProjectPath & "images/tags/ico_twitter2.png'"" alt=""Envia a twitter""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href='http://technorati.com/faves?add=" & strLink & "'><img src=""" & ProjectPath & "images/tags/ico_technorati2.png"" border=""0"" id=""technorati"" name=""technorati"" onmouseover=""document.technorati.src='" & ProjectPath & "images/tags/ico_technorati2.png'"" onmouseout=""document.technorati.src='" & ProjectPath & "images/tags/ico_technorati2.png'"" alt=""Envia a technorati""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href=""http://www.myspace.com/Modules/PostTo/Pages/?u=" & strLink & """ ><img src=""" & ProjectPath & "images/tags/3.jpg"" border=""0"" id=""myspace"" name=""myspace"" onmouseover=""document.myspace.src='" & ProjectPath & "images/tags/3.jpg'"" onmouseout=""document.myspace.src='" & ProjectPath & "images/tags/3.jpg'"" alt=""Envia a myspace""/></a>" & vbCrLf)

                sb.Append("		<a target='_blank' href=""https://favorites.live.com/quickadd.aspx?marklet=1&amp;mkt=en-us&amp;url=" & strLink & """><img src=""" & ProjectPath & "images/tags/8.jpg"" border=""0"" id=""live"" name=""live"" onmouseover=""document.live.src='" & ProjectPath & "images/tags/8.jpg'"" onmouseout=""document.live.src='" & ProjectPath & "images/tags/8.jpg'"" alt=""Envia a live""/></a>" & vbCrLf)

                'sb.Append("		<a target='_blank' href=""http://www.stumbleupon.com/submit?url=" & strLink & """><img src=""" & ProjectPath & "images/tags/9.jpg"" border=""0"" id=""stumbleupon"" name=""stumbleupon"" onmouseover=""document.stumbleupon.src='" & ProjectPath & "images/tags/9.jpg'"" onmouseout=""document.stumbleupon.src='" & ProjectPath & "images/tags/9.jpg'"" alt=""Envia a stumbleupon""/></a>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Public Function Rate(ByVal intArticleId As Integer, ByVal blIsForEmail As Boolean) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim arrValues As String() = Split("|Malo|Regular|Bueno|Muy Bueno|Excelente", "|")

            Dim intX As Integer = 0

            Dim intRating As Integer = 0

            Dim strLink As String = ""
            Dim strCachedVar As String = ""
            Dim strSessionName As String = ""

            Try

                strSessionName = "art" & intArticleId

                strCachedVar = "Rating" & intArticleId

                If IsNothing(Cache.Get(strCachedVar)) Then

                    Try
                        intRating = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_GetRatingByPageID " & intArticleId))
                    Catch ex As Exception
                    End Try

                    Cache.Insert(strCachedVar, intRating)

                End If

                intRating = Cache.Get(strCachedVar)

                sb.Append("<p align=""center"">")

                For intX = 1 To 5

                    If _objUser.getId > 0 Then

                        strLink = """javascript:"" onclick=""PostRating(" & intArticleId & ", " & intX & ", this);"""

                    Else

                        Session("RedirectURL") = QueryStringUrl()

                        strLink = ProjectPath & "logon.aspx"

                    End If



                    If intX <= intRating Then

                        If Session(strSessionName) <> "voted" And Not blIsForEmail Then

                            sb.Append("<a href=" & strLink & "><img src=" & ProjectPath & "images/stars/star_on.gif name=""star" & intX & """ border=""0"" onmouseover=""RateArticle(" & intX & ");"" onmouseout=""RateArticle(" & intRating & ");"" title=""" & arrValues(intX) & """></a>")

                        Else

                            sb.Append("<img src=" & ProjectPath & "images/stars/star_on.gif name=""star" & intX & """ border=""0"" >")

                        End If

                    Else

                        If Session(strSessionName) <> "voted" Then

                            sb.Append("<a href=" & strLink & "><img src=" & ProjectPath & "images/stars/star_off.gif name=""star" & intX & """ border=""0"" onmouseover=""RateArticle(" & intX & ");"" onmouseout=""RateArticle(" & intRating & ");""  title=""" & arrValues(intX) & """></a>")

                        Else

                            sb.Append("<img src=" & ProjectPath & "images/stars/star_off.gif name=""star" & intX & """ border=""0"" >")

                        End If

                    End If

                Next

                sb.Append("<div name=""rating"" id=""rating""></div>")

                sb.Append("</p>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function




        Public Function DisplayFlag() As String

            Dim strCountryCode As String = ""
            Dim strFlagPath As String = ""

            Try

                strCountryCode = Functions.GetCountryCode()

                Select Case strCountryCode

                    Case "", "--", " "

                        'do nothing

                    Case Else

                        strFlagPath = "<img src=" & ProjectPath & "images/flags/" & strCountryCode & ".gif border=0>"

                End Select

                Return strFlagPath

            Catch ex As Exception

                Throw ex

            Finally

            End Try

        End Function

        Public Sub DislpayBio(ByVal intUserId As Integer)

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim sb As New StringBuilder("")
            Dim arrFormValues() As String
            Dim strValues As String = ""
            Dim intIndex As Integer = 0

            '0 username Xavier Cabezas
            '1 email xavito@mesasub.com
            '2 city Chattanooga
            '3 state Tennessee
            '4 country Estados Unidos
            '5 picture xavito.jpg
            '6 otherinfo Xavier es un buen amigo de los Gulick
            '7 number of articles published by this user

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("sp_GetUserBio " & intUserId)

                If reader.HasRows Then

                    reader.Read()

                    For intIndex = 0 To 7

                        strValues = strValues & reader(intIndex) & "|"

                    Next

                    strValues = Left(strValues, Len(strValues) - 1)
                    arrFormValues = Split(strValues, "|")

                End If

                If Len(strValues) > 0 Then

                    sb.Append("<p align=center class=GEN>")

                    sb.Append(Functions.ShowPicture(intUserId, arrFormValues(5)))

                    sb.Append("</p>")


                    sb.Append("<p align=left class=GEN>")

                    sb.Append("Pa&#237;s: " & arrFormValues(4) & "<br>")

                    sb.Append("Ciudad: " & arrFormValues(2) & "<br>")

                    sb.Append("E-mail: <a href=mailto:" & arrFormValues(1) & ">" & arrFormValues(1) & "</a><br><br>")

                    sb.Append(Functions.ReplaceCR(arrFormValues(6)) & "<br><br>")

                    sb.Append(arrFormValues(0) & " ha colaborado con <a href=pages_by_user.aspx?user_id=" & intUserId & ">" & arrFormValues(7) & "</a> art&#237;culo(s).<br><br>")

                    If _objUser.getId = intUserId Then

                        sb.Append("<a href=" & _project_path & "registration.aspx>Edita/Cambia tu informaci&#243;n.</a><br><br>")

                    End If


                    sb.Append("</p>")


                    IsSimplePage = True
                    PageTitle = arrFormValues(0)
                    PageContent = sb.ToString()

                Else

                    PageContent = strValues

                End If


            Catch ex As Exception

                ShowError(ex)

            Finally

                reader.Close()
                reader = Nothing

                sb = Nothing

            End Try

        End Sub

#End Region

#Region "Subs"

        Sub New()

        End Sub

        Public Sub ShowError(ByVal excError As Exception)

            Dim objErrHandler As New ErrorHandler
            Dim sb As New StringBuilder("")

            Dim blDebugMode As Boolean = False

            Try

                Try
                    blDebugMode = CBool(System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode"))
                Catch ex As Exception
                End Try


                If excError.Message <> "Thread was being aborted." Then

                    If Not blDebugMode Then

                        Functions.SendMail(_support_account, _develop, "Error en Para Lideres", objErrHandler.ReturnHtmlErrorMessage(excError))

                    End If

                    Trace.Write(excError.TargetSite.Name)
                    Trace.Write(excError.Source)
                    Trace.Write(excError.Message)
                    Trace.Write(excError.ToString())

                    sb.Append("<img src=" & _project_path & "images/error.jpg alt=error align=left HSPACE=5 VSPACE=5>Hubo un error en la p&#225;gina que solicitaste.  Nuestro cuerpo t&#233;cnico ha sido informado al respecto.  Puedes regresar a la p&#225;gina anterior y tratar de nuevo.")

                    If _objUser.getSecurityLevel = 6 Then

                        sb.Append("<p><b>Informaci&#243;n para Adminstrador</b></p>")
                        sb.Append("<p><u>Error Information</u><p>" & excError.TargetSite.Name & "<p>" & excError.Source & "<p>" & excError.Message & "<p>" & excError.ToString())

                    End If

                    Me.PageTitle = "Error"
                    Me.PageContent = sb.ToString()

                End If

            Catch ex As Exception

                Trace.Warn(ex.ToString())

            Finally

                objErrHandler = Nothing
                sb = Nothing

            End Try

        End Sub


        Public Function UpdateCountryCodes() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim CountryName As String() = {"N/A", "Asia/Pacific Region", "Europe", "Andorra", "United Arab Emirates", "Afghanistan", "Antigua and Barbuda", "Anguilla", "Albania", "Armenia", "Netherlands Antilles", "Angola", "Antarctica", "Argentina", "American Samoa", "Austria", "Australia", "Aruba", "Azerbaijan", "Bosnia and Herzegovina", "Barbados", "Bangladesh", "Belgica", "Burkina Faso", "Bulgaria", "Bahrain", "Burundi", "Benin", "Bermuda", "Brunei Darussalam", "Bolivia", "Brasil", "Bahamas", "Bhutan", "Bouvet Island", "Botswana", "Belarus", "Belice", "Canada", "Cocos (Keeling) Islands", "Congo, The Democratic Republic of the", "Central African Republic", "Congo", "Switzerland", "Cote D'Ivoire", "Cook Islands", "Chile", "Cameroon", "China", "Colombia", "Costa Rica", "Cuba", "Cape Verde", "Christmas Island", "Cyprus", "Czzech Republic", "Alemania", "Djibouti", "Dinamarca", "Dominica", "Republica Dominicana", "Algeria", "Ecuador", "Estonia", "Egipto", "Western Sahara", "Eritrea", "España", "Ethiopia", "Finlandia", "Fiji", "Malvinas", "Micronesia, Federated States of", "Faroe Islands", "Francia", "Francia, Metropolitan", "Gabon", "Reino Unido", "Grenada", "Georgia", "French Guiana", "Ghana", "Gibraltar", "Greenland", "Gambia", "Guinea", "Guadeloupe", "Equatorial Guinea", "Greece", "South Georgia and the South Sandwich Islands", "Guatemala", "Guam", "Guinea-Bissau", "Guyana", "Hong Kong", "Heard Island and McDonald Islands", "Honduras", "Croatia", "Haiti", "Hungria", "Indonesia", "Irlanda", "Israel", "India", "British Indian Ocean Territory", "Iraq", "Iran, Islamic Republic of", "Iceland", "Italia", "Jamaica", "Jordan", "Japan", "Kenia", "Kyrgyzstan", "Cambodia", "Kiribati", "Comoros", "Saint Kitts and Nevis", "Korea, Democratic People's Republic of", "Korea, Republic of", "Kuwait", "Cayman Islands", "Kazakstan", "Lao People's Democratic Republic", "Lebanon", "Saint Lucia", "Liechtenstein", "Sri Lanka", "Liberia", "Lesotho", "Lithuania", "Luxembourg", "Latvia", "Libyan Arab Jamahiriya", "Maruecos", "Monaco", "Moldova, Republic of", "Madagascar", "Marshall Islands", "Macedonia, the Former Yugoslav Republic of", "Mali", "Myanmar", "Mongolia", "Macau", "Northern Mariana Islands", "Martinique", "Mauritania", "Montserrat", "Malta", "Mauritius", "Maldives", "Malawi", "Mexico", "Malaysia", "Mozambique", "Namibia", "New Caledonia", "Niger", "Norfolk Island", "Nigeria", "Nicaragua", "Netherlands", "Norway", "Nepal", "Nauru", "Niue", "New Zealand", "Oman", "Panama", "Peru", "French Polynesia", "Papua New Guinea", "Philippines", "Pakistan", "Poland", "Saint Pierre and Miquelon", "Pitcairn", "Puerto Rico", "Palestinian Territory, Occupied", "Portugal", "Palau", "Paraguay", "Qatar", "Reunion", "Romania", "Russian Federation", "Rwanda", "Saudi Arabia", "Solomon Islands", "Seychelles", "Sudan", "Sweden", "Singapore", "Saint Helena", "Slovenia", "Svalbard and Jan Mayen", "Slovakia", "Sierra Leone", "San Marino", "Senegal", "Somalia", "Suriname", "Sao Tome and Principe", "El Salvador", "Syrian Arab Republic", "Swaziland", "Turks and Caicos Islands", "Chad", "French Southern Territories", "Togo", "Thailand", "Tajikistan", "Tokelau", "Turkmenistan", "Tunisia", "Tonga", "East Timor", "Turkey", "Trinidad and Tobago", "Tuvalu", "Taiwan, Province of China", "Tanzania, United Republic of", "Ukraine", "Uganda", "United States Minor Outlying Islands", "Estados Unidos", "Uruguay", "Uzbekistan", "El Vaticano", "Saint Vincent and the Grenadines", "Venezuela", "Virgin Islands, British", "Virgin Islands, U.S.", "Vietnam", "Vanuatu", "Wallis and Futuna", "Samoa", "Yemen", "Mayotte", "Yugoslavia", "South Africa", "Zambia", "Zaire", "Zimbabwe", "Anonymous Proxy", "Satellite Provider"}
            Dim CountryCode As String() = {"--", "AP", "EU", "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AN", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AZ", "BA", "BB", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BM", "BN", "BO", "BR", "BS", "BT", "BV", "BW", "BY", "BZ", "CA", "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN", "CO", "CR", "CU", "CV", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EC", "EE", "EG", "EH", "ER", "ES", "ET", "FI", "FJ", "FK", "FM", "FO", "FR", "FX", "GA", "GB", "GD", "GE", "GF", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS", "GT", "GU", "GW", "GY", "HK", "HM", "HN", "HR", "HT", "HU", "ID", "IE", "IL", "IN", "IO", "IQ", "IR", "IS", "IT", "JM", "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ", "LA", "LB", "LC", "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "MG", "MH", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PM", "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RU", "RW", "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "ST", "SV", "SY", "SZ", "TC", "TD", "TF", "TG", "TH", "TJ", "TK", "TM", "TN", "TO", "TP", "TR", "TT", "TV", "TW", "TZ", "UA", "UG", "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI", "VN", "VU", "WF", "WS", "YE", "YT", "YU", "ZA", "ZM", "ZR", "ZW", "A1", "A2"}

            Dim intX As Integer = 0
            Dim intCountryId As Integer = 0
            Dim intColumnCount As Integer = 0

            Dim strSQL As String = ""

            Try

                Try

                    intColumnCount = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_GetTableStructureRowCount 'countries'"))

                Catch ex As Exception

                    Throw ex

                End Try


                If intColumnCount = 2 Then

                    strSQL = "ALTER TABLE countries ADD country_code varchar(4) NULL"

                    Try

                        ParaLideres.GenericDataHandler.ExecSQL(strSQL)

                    Catch ex As Exception

                        Throw ex

                    End Try

                End If

                For intX = 0 To UBound(CountryName)

                    reader = ParaLideres.GenericDataHandler.GetRecords("SELECT id FROM countries WHERE name = '" & Functions.FormatString(CountryName(intX)) & "'")

                    If reader.HasRows Then

                        reader.Read()

                        intCountryId = reader(0)

                        strSQL = "UPDATE countries SET country_code = '" & CountryCode(intX) & "' WHERE id = " & intCountryId

                        ParaLideres.GenericDataHandler.ExecSQL(strSQL)

                        sb.Append("</br>" & CountryName(intX) & " was found on db.")
                        'sb.Append("</br>" & strSQL)

                    Else

                        sb.Append("</br><font color=red>" & CountryName(intX) & "</font> was not found on db.")

                    End If

                Next

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing
                reader.Close()
                reader = Nothing

            End Try

        End Function

        'Protected Overrides Sub OnError(ByVal e As System.EventArgs)

        '    Dim LastError As Exception = Server.GetLastError()
        '    Dim objErrHandler As New ErrorHandler

        '    Try

        '        Functions.SendMail(_support_account, _support_account, "Error en Para Lideres", objErrHandler.ReturnHtmlErrorMessage(LastError))

        '        Trace.Warn(LastError.ToString())

        '        objErrHandler = Nothing
        '        LastError = Nothing

        '        Me.PageTitle = "ERROR"
        '        Me.PageContent = "Hubo un error en la p&#225;gina que solicitaste.  Nuestro cuerpo t&#233;cnico ha sido informado al respecto.  Puedes regresar a la p&#225;gina anterior y trata de nuevo."

        '    Catch ex As Exception

        '        'do nothing

        '    Finally

        '        LastError = Nothing
        '        objErrHandler = Nothing

        '    End Try

        'End Sub

#Region "Page Cycle"


        Protected Overrides Sub Construct()

            Trace.IsEnabled = System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode")

            Trace.Write("Start Construct")

            If IsNothing(ParaLideres.GenericDataHandler.ConnectionString) Then


                Trace.Warn("DataHandler.GenericDataHandler.ConnectionString: IsNothing")

            Else

                'Trace.Write("Info", "Connection string is not nothing")

                'Trace.Write("DataHandler.GenericDataHandler.ConnectionString: " & ParaLideres.GenericDataHandler.ConnectionString)

            End If

            Try

                EnableViewState = False

            Catch exc As Exception

                ShowError(exc)

            End Try

            Trace.Write("End Construct")

        End Sub

        Protected Overrides Sub OnInit(ByVal e As System.EventArgs)

            Dim intFoundBloqued As Integer = 0

            Session("IPAddress") = Request.UserHostAddress()

            Try
                _user_id = CInt(Session("user_id"))
            Catch ex As Exception
            End Try


            Try

                If Not ParaLideres.GenericDataHandler.TestConnection() Then

                    Response.Redirect(ProjectPath & "out_of_service.aspx")

                End If

                Try
                    intFoundBloqued = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_GetBloquedIPAddress '" & Session("IPAddress") & "'"))
                Catch ex As Exception
                End Try


                If intFoundBloqued > 0 Then Response.Redirect("http://" & Session("IPAddress"))


                _objUser = New DataLayer.reg_users(_user_id)

                If _objUser.getId = 1 Then Trace.IsEnabled = True

                _strScript = System.IO.Path.GetFileName(Me.Request.FilePath)

                Trace.Write(_strScript)

            Catch ex As Exception

            End Try

        End Sub

        'Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)

        '    'If Request.HttpMethod() = "POST" And InStr(Request.ServerVariables("URL"), "post") > 0 Then

        '    '    Trace.Write("Don't show the load screen")

        '    'Else

        '    '    Trace.Write("Show the load screen")

        '    'End If

        'End Sub

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

            If _is_under_construction Then

                _simple_page = True
                _intPageFormat = PageFormats.PrintFormat
                _page_title = "Fuera De Servicio"
                _page_content = "Por ahora estamos fuera de servicio, por favor regresa en unos pocos minutos mas."

                writer.Write(PageTemplate2())

            Else 'If _is_under_construction Then

                If Session("test") = "oui" And Not _blRequiresLogin Then

                    Dim objNewDesign As ParaLideres.Design2010

                    Try

                        objNewDesign = New ParaLideres.Design2010

                        writer.Write(objNewDesign.GenerateContent(_objUser, _page_title, _page_content, PageFormats.NormalPage))

                    Catch ex As Exception

                        Throw ex

                    Finally

                        objNewDesign = Nothing

                    End Try


                Else 'If Session("test") = "oui" And Not _blRequiresLogin Then

                    If _blRequiresLogin And _objUser.getId > 0 Then

                        writer.Write(PageTemplate2())

                    ElseIf Not _blRequiresLogin Then

                        writer.Write(PageTemplate2())

                    Else

                        Session("RedirectURL") = QueryStringUrl()

                        Response.Redirect(ProjectPath & "logon.aspx")

                    End If

                End If 'If Session("test") = "oui" And Not _blRequiresLogin Then


            End If 'If _is_under_construction Then



        End Sub

        Protected Overrides Sub OnUnload(ByVal e As System.EventArgs)

            Trace.Write("OnUnload called!")

            _objUser = Nothing

        End Sub

        Public Overrides Sub Dispose()


            MyBase.Dispose()

        End Sub

#End Region

#End Region

#Region "Security"

        Public Sub ShowNoAuthorized()

            _page_title = "Access Denied"
            _page_content = "You have no authorization to view or access this page.<br>If you think that this is a mistake please contact support."

        End Sub

#End Region





    End Class


End Namespace
