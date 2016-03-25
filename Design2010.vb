Imports System.Web
Imports System.Text

Namespace Paralideres

    Public Class Design2010


        Private _strProjectPath As String = "http://" & HttpContext.Current.Request.ServerVariables("HTTP_HOST") & "/"

        Public Function GenerateContent(ByVal objUser As DataLayer.reg_users, ByVal strPageTitle As String, ByVal strPageContent As String, ByVal intPageFormat As PageTemplate.PageFormats) As String

            Select Case intPageFormat

                Case PageTemplate.PageFormats.NormalPage


                    Return PageTemplateNormal(objUser, strPageTitle, strPageContent)

                    'TODO: more options

            End Select


        End Function


        Private Function PageTemplateNormal(ByVal objUser As DataLayer.reg_users, ByVal strPageTitle As String, ByVal strPageContent As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf)

                sb.Append("<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf)

                sb.Append(PageHead(strPageTitle))

                sb.Append("<body>" & vbCrLf)

                'TODO: Show line above only on home page
                'sb.Append(CreateNewUser())

                sb.Append(MainWraper(objUser, strPageTitle, strPageContent))

                sb.Append("</body>" & vbCrLf)

                sb.Append("</html>" & vbCrLf)


                Return sb.ToString()

            Catch ex As Exception

                HttpContext.Current.Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function CreateNewUser() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("	<div id=""dialog-form"" title=""Create new user"" style=""display:none;"">" & vbCrLf)

                sb.Append("        <p class=""validateTips"">All form fields are required.</p>" & vbCrLf)

                sb.Append("        <form>" & vbCrLf)

                sb.Append("        <fieldset>" & vbCrLf)

                sb.Append("            <label for=""name"">Name</label>" & vbCrLf)

                sb.Append("            <input type=""text"" name=""name"" id=""name"" class=""text ui-widget-content"" />" & vbCrLf)

                sb.Append("            <label for=""email"">Email</label>" & vbCrLf)

                sb.Append("            <input type=""text"" name=""email"" id=""email"" value="""" class=""text ui-widget-content"" />" & vbCrLf)

                sb.Append("            <label for=""password"">Password</label>" & vbCrLf)

                sb.Append("            <input type=""password"" name=""password"" id=""password"" value="""" class=""text ui-widget-content"" />" & vbCrLf)

                sb.Append("        </fieldset>" & vbCrLf)

                sb.Append("        </form>" & vbCrLf)

                sb.Append("	</div>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Private Function TopUserBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("        	<div id=""top_user_box"">" & vbCrLf)

                sb.Append("            	<div id=""user_box"">" & vbCrLf)

                sb.Append("            	    <a href=""/publish_my_materials.aspx"" class=""link_btn_round_md_subir"">SUBIR</a>" & vbCrLf)

                sb.Append("            	    <a href=""/registration.aspx"" id=""userbox_reg_link""class=""userbox_reg_link"" >REGISTRATE</a>" & vbCrLf)

                sb.Append("            	    <a href=""/logon.aspx"" class=""userbox_log_link"">ENTRAR</a>" & vbCrLf)

                sb.Append("            	</div>" & vbCrLf)

                sb.Append("          </div>" & vbCrLf)


                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function MainWraper(ByVal objUser As DataLayer.reg_users, ByVal strPageTitle As String, ByVal strPageContent As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("	<div id=""main_wraper"" class=""clearfix"">" & vbCrLf)


                sb.Append("    	<!-- Starts Header-->" & vbCrLf)

                sb.Append("    	<div id=""header"">" & vbCrLf)

                sb.Append("        	<!-- Top Navigation bar // Login, Register, Follow etc.. -->" & vbCrLf)

                sb.Append(TopUserBox())

                sb.Append("            <!-- Starts Main Header Box -->" & vbCrLf)

                sb.Append(MainHeadBox())

                sb.Append("        <!-- ends Main header Box -->" & vbCrLf)

                sb.Append("     </div>" & vbCrLf)

                sb.Append("        <!-- ends Header-->" & vbCrLf)

                sb.Append("        <!-- Starts Content -->" & vbCrLf)

                'TODO:
                sb.Append(Content(objUser, strPageTitle, strPageContent))

                sb.Append("        <!-- Ends Content -->" & vbCrLf)

                sb.Append(Footer())

                sb.Append(" </div>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function MainHeadBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("            <div id=""main_head_box"">" & vbCrLf)

                sb.Append("            	<!-- Logo Box -->" & vbCrLf)

                sb.Append("            	<div class=""logo_box""><a href=""/home.aspx""><img src=""" & _strProjectPath & "assets/imgs/gral/paralideres_logo.png"" alt=""ParaLideres.org | Recursos - Capacitaci&oacute;n - Comunidad"" title=""ParaLideres.org | Recursos - Capacitaci&oacute;n - Comunidad"" /></a></div>" & vbCrLf)

                sb.Append("            <!-- Starts Menu-->" & vbCrLf)

                sb.Append("            <div id=""menu"">" & vbCrLf)

                sb.Append("                <!-- Search Menu UL Bar -->" & vbCrLf)

                sb.Append("                <div id=""menu_bar"">" & vbCrLf)

                sb.Append("                    <ul>" & vbCrLf)

                sb.Append("                        <li><a href=""javascript:ShowAjaxByDiv('menu_display_box', '" & _strProjectPath & "show_submenu.aspx?parent_id=21');"" id=""rec_btn"">RECURSOS</a></li>" & vbCrLf)

                sb.Append("                        <li><a href=""javascript:ShowAjaxByDiv('menu_display_box', '" & _strProjectPath & "show_submenu.aspx?parent_id=20');"">CAPACITACI&Oacute;N</a></li>" & vbCrLf)

                sb.Append("                        <li><a href=""http://paralideresblog.blogspot.com"" target=""new"">COMUNIDAD</a></li>" & vbCrLf)

                sb.Append("                    </ul>" & vbCrLf)

                sb.Append("                </div>" & vbCrLf)

                sb.Append("                <!-- ends Menu Bar -->" & vbCrLf)

                sb.Append("          </div>" & vbCrLf)

                sb.Append("        <!-- Ends Menu --> " & vbCrLf)

                sb.Append(SearchBox())

                sb.Append("</div>")

                'Hidden Top Menu Options
                sb.Append("         <div id=""menu_display_box"">" & vbCrLf)

                sb.Append("         </div>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function Content(ByVal objUser As DataLayer.reg_users, ByVal strPageTitle As String, ByVal strPageContent As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("        <!-- Starts Content -->" & vbCrLf)

                sb.Append("        <div id=""content"" class=""clearfix"">" & vbCrLf)

                If strPageTitle <> "" Then sb.Append(strPageTitle & "<br/>")

                sb.Append(strPageContent)

                sb.Append("            <!-- Starts Sidebar -->	" & vbCrLf)

                sb.Append(SideBarBox())

                sb.Append("            <!-- ends SideBar -->" & vbCrLf)

                sb.Append("        </div>" & vbCrLf)

                sb.Append("    <!-- Ends Content -->" & vbCrLf)


                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function HomePageContent() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("        	<!-- Main News Announcer -->" & vbCrLf)

                sb.Append(MainAnnouncerBox())

                sb.Append("            <!-- ends Main News Announcer -->" & vbCrLf)

                sb.Append("            <!-- Central Content box -->" & vbCrLf)

                sb.Append(CentralContentBox())

                sb.Append("            <!-- ends Central Content box -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function




        Private Function MainAnnouncerBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("            <div class=""main_announcer_box"">" & vbCrLf)

                sb.Append("                <p><span class=""ma_title"">Bienvenido!</span><br />" & vbCrLf)

                sb.Append("                Durante los últimos 11 años hemos servido como una herramienta de difusión y distribución de más de <em>25.000</em> materiales y recursos para el trabajo con Jóvenes y Adolescentes. Ahora queremos invitarte para que puedas hacer uso de las facilidades que ParaLideres tienes para tí.<br /><br />" & vbCrLf)

                sb.Append("                </p>" & vbCrLf)

                sb.Append("                <div class=""registry_notice"">" & vbCrLf)

                sb.Append("                	<a href=""/registration.aspx""><img src=""" & _strProjectPath & "assets/imgs/gral/btn_round_bg_join_bkg.png"" alt=""unete hoy"" /></a>" & vbCrLf)

                sb.Append("               </div>" & vbCrLf)

                sb.Append("            </div>" & vbCrLf)

                sb.Append("            <!-- ends Main News Announcer -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function CentralContentBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("            <!-- Central Content box -->" & vbCrLf)

                sb.Append("            <div id=""central_content"">" & vbCrLf)


                sb.Append("            	<!-- Starts last_updates widget -->" & vbCrLf)

                sb.Append(LastUpdates())

                sb.Append("                <!-- Ends Last_updates widget-->" & vbCrLf)

                sb.Append("                <!-- Starts Stream's Box -->" & vbCrLf)

                sb.Append(NewsStreamBox())

                sb.Append("                <!-- ends Stream's Widget Boz -->" & vbCrLf)

                sb.Append("                <!-- Starts Author's Widget box -->" & vbCrLf)

                sb.Append(Autores())

                sb.Append("                <!-- ends Author's Widget Box -->" & vbCrLf)

                sb.Append("            </div>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function LastUpdates() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strPath As String = ""

            Try

                sb.Append("            	<!-- Starts last_updates widget -->" & vbCrLf)

                sb.Append("                <div id=""last_updates"">" & vbCrLf)


                sb.Append("                    <div id=""tabs"">" & vbCrLf)

                sb.Append("                        <ul class=""tabs-nav"">" & vbCrLf)

                strPath = _strProjectPath & "featured.aspx?showlist=1"

                sb.Append("                            <li class=""default""><a href=""javascript:ShowAjaxByDiv('DivLastUpdates','" & strPath & "');"">Lo Último</a></li>" & vbCrLf)

                strPath = _strProjectPath & "featured.aspx?showlist=2"

                sb.Append("                            <li class=""default""><a href=""javascript:ShowAjaxByDiv('DivLastUpdates','" & strPath & "');"">Destacado</a></li>" & vbCrLf)

                strPath = _strProjectPath & "featured.aspx?showlist=3"

                sb.Append("                            <li class=""default""><a href=""javascript:ShowAjaxByDiv('DivLastUpdates','" & strPath & "');"">Popular</a></li>" & vbCrLf)

                sb.Append("                        </ul>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)

                sb.Append("                    <div id=""DivLastUpdates"">" & vbCrLf) 'class=""tabs_more""

                sb.Append(Paralideres.Functions.ShowFeatured(1))

                sb.Append("                    </div>" & vbCrLf)

                sb.Append("                </div>" & vbCrLf)

                sb.Append("                <!-- Ends Last_updates widget-->" & vbCrLf)






                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        'Private Function Tabs() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try
        '        sb.Append("                    <div id=""tabs"">" & vbCrLf)

        '        sb.Append("                        <ul>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#tabs-1"">Lo Último</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#tabs-2"">Destacado</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#tabs-3"">Popular</a></li>" & vbCrLf)

        '        sb.Append("                        </ul>" & vbCrLf)


        '        sb.Append("                        <div id=""tabs-1"">" & vbCrLf)

        '        sb.Append("                        	<div>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Rompehielos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Rompehielos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Dinámicas/Juegos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">Lorem ipsum dolor sit amet, consectetur...</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Alejandroino...</a> el 21/03/09</span>                            </div>" & vbCrLf)

        '        sb.Append("                           <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                          		<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Articulos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                           </div>" & vbCrLf)


        '        sb.Append("                        </div>" & vbCrLf)


        '        sb.Append("                        <div id=""tabs-2"">" & vbCrLf)

        '        sb.Append("                            Cargando..." & vbCrLf)

        '        sb.Append("                        </div>" & vbCrLf)


        '        sb.Append("                        <div id=""tabs-3"">" & vbCrLf)

        '        sb.Append("                            Cargando..." & vbCrLf)

        '        sb.Append("                        </div>" & vbCrLf)


        '        sb.Append("                    </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function


        'Private Function Tabs1() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try
        '        sb.Append("                        <div id=""tabs-1"">" & vbCrLf)

        '        sb.Append("                        	<div>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Rompehielos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Rompehielos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)


        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Dinámicas/Juegos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">Lorem ipsum dolor sit amet, consectetur...</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Alejandroino...</a> el 21/03/09</span>                            </div>" & vbCrLf)

        '        sb.Append("                           <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                          		<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Articulos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                           </div>" & vbCrLf)


        '        sb.Append("                        </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function


        'Private Function LuContentBox() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try
        '        sb.Append("                            <div class=""lu_content_box"">" & vbCrLf)

        '        sb.Append("                            	<a href=""" & _strProjectPath & "#"" class=""link_btn_round_sm"">Rompehielos</a>" & vbCrLf)

        '        sb.Append("                            	<img class=""lu_doc_icon"" src=""" & _strProjectPath & "assets/imgs/doc_type_icons/doc.png"" />" & vbCrLf)

        '        sb.Append("                                <h1><a href=""" & _strProjectPath & "#"">El Evangelio segun los simpsons</a></h1>" & vbCrLf)

        '        sb.Append("                                <span>Por <a href=""" & _strProjectPath & "#"">Flavio Calvo</a> el 21/03/09</span>" & vbCrLf)

        '        sb.Append("                            </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function


        Private Function NewsStreamBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim objQuery As New Google.GData.Client.FeedQuery()
            Dim objFeed As Google.GData.Client.AtomFeed
            Dim objService As Google.GData.Client.Service
            Dim objEntry As Google.GData.Client.AtomEntry

            Dim strBlogId As String = "3403368607335455694"

            Try
                sb.Append("                <!-- Starts Stream's Box -->" & vbCrLf)

                sb.Append("                <div id=""news_stream_box"" class=""clearfix"">" & vbCrLf)


                sb.Append("                    <div class=""nws_header"">" & vbCrLf)

                sb.Append("                    	<h1>BLOG STREAM</h1>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                'sb.Append("                    <div class=""nws_option_selector"">" & vbCrLf)

                'sb.Append("                	<select name=""nw_select_tool"">" & vbCrLf)

                'sb.Append("                    	<option>Recientes</option>" & vbCrLf)

                'sb.Append("                        <option>Por Autor</option>" & vbCrLf)

                'sb.Append("                        <option>Mas Vistos</option>" & vbCrLf)

                'sb.Append("                    </select>" & vbCrLf)

                'sb.Append("                    <a href=""" & _strProjectPath & "#"" class=""nws_suscribe"">SUSCR&Iacute;BETE</a>" & vbCrLf)

                'sb.Append("                    <a class=""nws_more"" href=""" & _strProjectPath & "#"">VER M&Aacute;S</a>")

                'sb.Append("                    </div>" & vbCrLf)


                sb.Append("                    <div class=""nws_content"">" & vbCrLf)

                sb.Append("                    	<ul>" & vbCrLf)


                objQuery.Uri = New Uri("http://www.blogger.com/feeds/" & strBlogId & "/posts/default")

                objQuery.MinPublication = Date.Today.AddDays(-20)

                objQuery.MaxPublication = Date.Today

                objQuery.NumberToRetrieve = 5

                objService = New Google.GData.Client.Service

                objFeed = objService.Query(objQuery)

                For Each objEntry In objFeed.Entries

                    sb.Append("                        <!-- stream -->" & vbCrLf)

                    sb.Append("                        <li>" & vbCrLf)

                    sb.Append("                        	<div class=""nws_stream"">" & vbCrLf)

                    sb.Append("                                <img src=""" & _strProjectPath & "assets/imgs/authors/blogger.gif"" />" & vbCrLf)

                    sb.Append("                                <div>" & vbCrLf)

                    sb.Append("                                     <h1><a href=""" & objEntry.AlternateUri.Content & """ target=""new"">" & objEntry.Title.Text & "</a></h1>" & vbCrLf)

                    If objEntry.Summary.Text <> "" Then sb.Append("                                     <br />" & objEntry.Summary.Text & "<br />" & vbCrLf)

                    sb.Append("                                     <span>Autor:" & objEntry.Authors.Item(0).Name & " | " & objEntry.Published & "</span>")

                    sb.Append("                                </div>" & vbCrLf)

                    sb.Append("                            </div>" & vbCrLf)

                    sb.Append("                        </li>" & vbCrLf)

                    sb.Append("                        <!-- end stream -->" & vbCrLf)

                    'sb.Append("<br /><br />")
                    'sb.Append("<br />Entry ID: " & objEntry.Id.ToString)
                    'sb.Append("<br />Entry Title: " & objEntry.Title.Text)
                    'sb.Append("<br />Entry Published On: " & objEntry.Published)
                    'sb.Append("<br />Entry Last Updated: " & objEntry.Updated)
                    'sb.Append("<br />Entry Summary: " & objEntry.Summary.Text)
                    'sb.Append("<br />Entry Author: " & objEntry.Authors.Item(0).Name)
                    'sb.Append("<br />objEntry.AlternateUri.Content: " & objEntry.AlternateUri.Content)

                Next

                sb.Append("                        </ul>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)

                sb.Append("                </div>" & vbCrLf)

                sb.Append("                <!-- ends Stream's Widget Boz -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

                objQuery = Nothing

                objService = Nothing

                objFeed = Nothing

                objEntry = Nothing

            End Try

        End Function


        Private Function NwsHeader() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("                    <div class=""nws_header"">" & vbCrLf)

                sb.Append("                    	<h1>BLOG STREAM</h1>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Private Function Autores() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("                <!-- Starts Author's Widget box -->" & vbCrLf)

                sb.Append("                <div id=""content_stream_box"" class=""clearfix"">" & vbCrLf)


                sb.Append("                	<div class=""cnt_header"">" & vbCrLf)

                sb.Append("                    	<h1>CONTENIDO</h1>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                sb.Append("                    <div class=""cnt_option_selector"">" & vbCrLf)

                sb.Append("                	    <select name=""cnt_select_tool"" id=""cnt_select_tool"" onchange=""ShowAjaxByDiv('autores', '/autores.aspx?selection=' + document.getElementById('cnt_select_tool').value);"">" & vbCrLf)  'ShowAjaxByDiv(strDivName, strURL)  'onchange=""alert(document.getElementById('cnt_select_tool').value);""

                sb.Append("                        <option value=""1"">Recientes</option>" & vbCrLf)

                sb.Append("                        <option  value=""2"">Mas Vistos</option>" & vbCrLf)

                sb.Append("                     </select>" & vbCrLf)

                'sb.Append("                    <a class=""cnt_more"" href=""" & Functions.ProjectPath & "#"">VER M&Aacute;S</a>")

                sb.Append("                     </div>" & vbCrLf)

                sb.Append("                    <div name=""autores"" id=""autores"" class=""cnt_content"">" & vbCrLf)

                sb.Append(AutoresContent(1))

                sb.Append("                    </div>" & vbCrLf)


                sb.Append("                </div>" & vbCrLf)

                sb.Append("                <!-- ends Author's Widget Box -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Public Shared Function AutoresContent(ByVal intSelection As Integer) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim strAuthorName As String = ""
            Dim strFileName As String = ""
            Dim strPicName As String = ""
            Dim strPath As String = HttpContext.Current.Request.PhysicalApplicationPath & "\files\"
            Dim strPathAlt As String = HttpContext.Current.Request.PhysicalApplicationPath & "\images\"
            Dim strSQL As String = ""

            Dim intAuthorId As Integer = 0
            Dim intSexo As Integer = 0

            Try

                If intSelection = 1 Then

                    strSQL = "proc_GetAuthorsForHomePage"

                Else

                    strSQL = "proc_GetAuthorsForHomePageByRating"

                End If


                reader = Paralideres.GenericDataHandler.GetRecords(strSQL)

                If reader.HasRows Then

                    sb.Append("                    	<ul>" & vbCrLf)

                    Do While reader.Read

                        intAuthorId = reader(0)

                        strAuthorName = UCase(reader(1))

                        strPicName = reader(2)

                        intSexo = reader(3)

                        sb.Append("                        <!-- stream -->" & vbCrLf)

                        sb.Append("                        <li>" & vbCrLf)

                        sb.Append("                        	<div class=""cnt_stream"">" & vbCrLf)

                        sb.Append(Functions.ShowPicture(intAuthorId, strPicName, intSexo, "r_author_pic"))

                        'strFileName = "usr_" & intAuthorId & ".jpg"

                        'If System.IO.File.Exists(strPath & strFileName) Then

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "files/" & strFileName & """ width=""60""/>" & vbCrLf)

                        'ElseIf System.IO.File.Exists(strPath & strPicName) Then

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "files/" & strPicName & """ width=""60""/>" & vbCrLf)

                        'ElseIf System.IO.File.Exists(strPathAlt & strPicName) Then

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "images/" & strPicName & """ width=""60""/>" & vbCrLf)

                        'ElseIf intSexo = 1 Then

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "images/AvatarMasculino.jpg"" width=""60""/>" & vbCrLf)

                        'ElseIf intSexo = 2 Then

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "images/AvatarFemenino.jpg"" width=""60""/>" & vbCrLf)

                        'Else

                        '    sb.Append("                                <img src=""" & Functions.ProjectPath & "images/AvatarMasculino.jpg"" width=""60""/>" & vbCrLf)

                        'End If

                        sb.Append("                         <div>" & vbCrLf)

                        sb.Append("                                <h1>" & strAuthorName & "</h1>" & vbCrLf)

                        'sb.Append("                                <p>Miembro desde 2001</p>" & vbCrLf)

                        sb.Append("                                <span><a href=""" & Functions.ProjectPath & "bio.aspx?user_id=" & intAuthorId & """>PERF&Iacute;L</a>   <a href=""" & Functions.ProjectPath & "pages_by_user.aspx?user_id=" & intAuthorId & """>CONTRIBUCIONES</a></span></div>" & vbCrLf)

                        sb.Append("                            </div>" & vbCrLf)

                        sb.Append("                        </li>" & vbCrLf)

                        sb.Append("                        <!-- end stream -->" & vbCrLf)

                    Loop

                    sb.Append("                        </ul>" & vbCrLf)

                End If 'If reader.hasrows

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

                reader.Close()
                reader = Nothing


            End Try

        End Function


        Private Function SideBarBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("            <!-- Starts Sidebar -->	" & vbCrLf)

                sb.Append("            <div id=""sidebar_box"">" & vbCrLf)


                'ENCUESTA
                sb.Append("                <div class=""sb_box"">" & vbCrLf)


                sb.Append("                	<div class=""sb_title"">ENCUESTA</div>" & vbCrLf)

                sb.Append(Encuesta())

                sb.Append("                </div>" & vbCrLf)


                ''NUBE DE ETIQUETAS
                'sb.Append("                <div class=""sb_box"">" & vbCrLf)

                'sb.Append("                	<div class=""sb_title"">NUBE DE ETIQUETAS</div>" & vbCrLf)

                'sb.Append("                    <div class=""sb_content""></div>" & vbCrLf)

                'sb.Append("                </div>" & vbCrLf)


                sb.Append("            </div>" & vbCrLf)

                sb.Append("            <!-- ends SideBar -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function Footer() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try
                sb.Append("        <!-- starts Footer -->" & vbCrLf)

                sb.Append("        <div id=""footer"" class=""clearfix"">" & vbCrLf)

                sb.Append("        	<!-- Connections Box // Twitter Facebook etc -->" & vbCrLf)


                sb.Append("        	<div id=""connect_box"">" & vbCrLf)


                sb.Append("                <div class=""ft_title"">" & vbCrLf)

                sb.Append("                CON&Eacute;CTATE CON NOSOTROS EN L&Iacute;NEA" & vbCrLf)

                sb.Append("                </div>" & vbCrLf)


                sb.Append("                <div class=""ft_content"">" & vbCrLf)

                sb.Append("                <p class=""ft_social""><a href=""http://www.twitter.com/paralideres"" target=""new""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_twitter_btn.jpg"" /></a>" & vbCrLf)

                sb.Append("                <a href=""http://www.facebook.com/pages/Paralideres/290701650279"" target=""new""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_facebook_btn.jpg"" /></a>" & vbCrLf)

                sb.Append("                <a href=""http://paralideresblog.blogspot.com"" target=""new""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_rss_btn.jpg"" /></a>" & vbCrLf)

                'sb.Append("                <a href=""" & _strProjectPath & "#""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_delicious_btn.jpg"" /></a></p>" & vbCrLf)

                sb.Append("                <p class=""ft_aditional_links"">" & vbCrLf)

                sb.Append("                <a href=""/comments.aspx"">Contáctenos</a><BR />" & vbCrLf)

                sb.Append("                <a href=""/article.aspx?page_id=652"">Qué Creemos</a>" & vbCrLf)

                sb.Append("                </p>" & vbCrLf)

                sb.Append("                </div>" & vbCrLf)


                sb.Append("            </div>" & vbCrLf)

                sb.Append("            <!-- Navigation Box //  -->" & vbCrLf)


                sb.Append("            <div id=""nav_box"" class=""clearfix"">" & vbCrLf)


                sb.Append("            	<div class=""ft_title"">" & vbCrLf)

                sb.Append("	            PUEDES VISITAR TAMBIEN ..." & vbCrLf)

                sb.Append("                </div>" & vbCrLf)


                sb.Append("                <div class=""ft_content"">" & vbCrLf)

                sb.Append("                	<div class=""ft_sitemap_col clearfix"">" & vbCrLf)

                sb.Append("                        <ul>" & vbCrLf)

                sb.Append("                        	<li>RECURSOS</li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                        </ul>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                sb.Append("                    <div class=""ft_sitemap_col clearfix"">" & vbCrLf)

                sb.Append("                        <ul>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                        </ul>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                sb.Append("                    <div class=""ft_sitemap_col clearfix"">" & vbCrLf)

                sb.Append("                        <ul>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

                sb.Append("                        </ul>" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                sb.Append("                </div>" & vbCrLf)


                sb.Append("            </div>" & vbCrLf)


                sb.Append("            <div id=""copy_box"">ParaLideres.org &copy; 2010 | <a href=""" & _strProjectPath & "#1"">T&eacute;rminos</a> | <a href=""" & _strProjectPath & "#1"">Privacidad</a> | Todos los derechos reservados </div>" & vbCrLf)

                sb.Append("        </div>" & vbCrLf)

                sb.Append("        <!-- ends Footer Box -->" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        'Private Function ConnectBox() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try
        '        sb.Append("        	<div id=""connect_box"">" & vbCrLf)


        '        sb.Append("                <div class=""ft_title"">" & vbCrLf)

        '        sb.Append("                CON&Eacute;CTATE CON NOSOTROS EN L&Iacute;NEA" & vbCrLf)

        '        sb.Append("                </div>" & vbCrLf)


        '        sb.Append("                <div class=""ft_content"">" & vbCrLf)

        '        sb.Append("                <p class=""ft_social""><a href=""" & _strProjectPath & "#""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_twitter_btn.jpg"" /></a>" & vbCrLf)

        '        sb.Append("                <a href=""" & _strProjectPath & "#""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_facebook_btn.jpg"" /></a>" & vbCrLf)

        '        sb.Append("                <a href=""" & _strProjectPath & "#""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_rss_btn.jpg"" /></a>" & vbCrLf)

        '        sb.Append("                <a href=""" & _strProjectPath & "#""><img title=""Twitter"" src=""" & _strProjectPath & "assets/imgs/gral/foot_delicious_btn.jpg"" /></a></p>" & vbCrLf)

        '        sb.Append("                <p class=""ft_aditional_links"">" & vbCrLf)

        '        sb.Append("                <a href=""" & _strProjectPath & "#"">Contáctenos</a><BR />" & vbCrLf)

        '        sb.Append("                <a href=""" & _strProjectPath & "#"">Qué Creemos</a>" & vbCrLf)

        '        sb.Append("                </p>" & vbCrLf)

        '        sb.Append("                </div>" & vbCrLf)


        '        sb.Append("            </div>" & vbCrLf)

        '        sb.Append("            <!-- Navigation Box //  -->" & vbCrLf)


        '        sb.Append("            <div id=""nav_box"" class=""clearfix"">" & vbCrLf)


        '        sb.Append("            	<div class=""ft_title"">" & vbCrLf)

        '        sb.Append("	            PUEDES VISITAR TAMBIEN ..." & vbCrLf)

        '        sb.Append("                </div>" & vbCrLf)


        '        sb.Append("                <div class=""ft_content"">" & vbCrLf)

        '        sb.Append("                	<div class=""ft_sitemap_col clearfix"">" & vbCrLf)

        '        sb.Append("                        <ul>" & vbCrLf)

        '        sb.Append("                        	<li>RECURSOS</li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                        </ul>" & vbCrLf)

        '        sb.Append("                    </div>" & vbCrLf)


        '        sb.Append("                    <div class=""ft_sitemap_col clearfix"">" & vbCrLf)

        '        sb.Append("                        <ul>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                        </ul>" & vbCrLf)

        '        sb.Append("                    </div>" & vbCrLf)


        '        sb.Append("                    <div class=""ft_sitemap_col clearfix"">" & vbCrLf)

        '        sb.Append("                        <ul>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                            <li><a href=""" & _strProjectPath & "#1"">Link</a></li>" & vbCrLf)

        '        sb.Append("                        </ul>" & vbCrLf)

        '        sb.Append("                    </div>" & vbCrLf)


        '        sb.Append("                </div>" & vbCrLf)


        '        sb.Append("            </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function


        Private Function PageHead(ByVal strPageTitle As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("<head>" & vbCrLf)

                sb.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbCrLf)

                sb.Append("<title>ParaLideres.org | " & strPageTitle & "</title>" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/main_style.css""  />" & vbCrLf)

                sb.Append("<link type=""text/css"" href=""" & _strProjectPath & "assets/css/pl_theme/jquery-ui-1.8.custom.css"" rel=""Stylesheet"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/menu_style.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/resource_style.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/last_updates.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/widgets/last_updates.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/widgets/survey.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/widgets/news_stream.css"" />" & vbCrLf)

                sb.Append("<link type=""text/css"" media=""screen"" rel=""stylesheet"" href=""" & _strProjectPath & "assets/css/widgets/content_stream.css"" />	" & vbCrLf)

                sb.Append("<script src=""" & _strProjectPath & "javascript/ajax.js"" type=""text/javascript""></script>" & vbCrLf)

                'sb.Append("<script type=""text/javascript"" src=""http://ajax.googleapis.com/ajax/libs/jqueryui/1.8.0/jquery-ui.min.js""></script>" & vbCrLf)

                'sb.Append("<script src=""" & _strProjectPath & "assets/js/init.js"" type=""text/javascript"" language=""javascript""></script>" & vbCrLf)

                sb.Append("<!--[if lte IE 6]>" & vbCrLf)

                sb.Append("<link rel=""stylesheet"" type=""text/css"" href=""" & _strProjectPath & "assets/css/png_ie6.css"" />" & vbCrLf)

                sb.Append("<![endif]-->" & vbCrLf)


                sb.Append("</head>" & vbCrLf)


                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        'Public Function UserBox() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try

        '        sb.Append("<div id=""user_box""><a href=""" & _strProjectPath & "#"" class=""link_btn_round_md_subir"">SUBIR</a><a id=""userbox_reg_link""class=""userbox_reg_link"" href=""" & _strProjectPath & "#"">REGISTRATE</a><a class=""userbox_log_link"" href=""" & _strProjectPath & "#"">ENTRAR</a></div>" & vbCrLf)


        '        Return sb.ToString()

        '    Catch ex As Exception

        '        HttpContext.Current.Trace.Warn(ex.ToString())

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function
        'Public Function LogoBox() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try

        '        sb.Append("<div class=""logo_box""><a href=""" & _strProjectPath & "#""><img src=""" & _strProjectPath & "assets/imgs/gral/paralideres_logo.png"" alt=""ParaLideres.org | Recursos - Capacitaci&oacute;n - Comunidad"" title=""ParaLideres.org | Recursos - Capacitaci&oacute;n - Comunidad"" /></a></div>" & vbCrLf)


        '        Return sb.ToString()

        '    Catch ex As Exception

        '        HttpContext.Current.Trace.Warn(ex.ToString())

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function

        'Private Function MenuBar() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try
        '        sb.Append("                <div id=""menu_bar"">" & vbCrLf)

        '        sb.Append("                    <ul>" & vbCrLf)

        '        sb.Append("                        <li><a href=""" & _strProjectPath & "#"" id=""rec_btn"">RECURSOS</a></li>" & vbCrLf)

        '        sb.Append("                        <li><a href=""" & _strProjectPath & "#"" id=""cap_btn"">CAPACITACI&Oacute;N</a></li>" & vbCrLf)

        '        sb.Append("                        <li><a href=""http://paralideresblog.blogspot.com"" target=""new"" id=""com_btn"">COMUNIDAD</a></li>" & vbCrLf)

        '        sb.Append("                    </ul>" & vbCrLf)

        '        sb.Append("                </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function

        Public Function SearchBox() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strFormName As String = "search_form"
            Dim strField As String = "shearchparam"

            Try

                sb.Append("             <div id=""search_box"">" & vbCrLf)

                sb.Append("                <form name=""search_form"" id=""search_form"" method=""post"" action=""" & Functions.ProjectPath & "search.aspx"" >" & vbCrLf)

                sb.Append("                    <input type=""text"" name=""" & strField & """ id=""" & strField & """ size=""40"" class=""search_txt_input"" />" & vbCrLf)

                sb.Append("                    <input type=""button""  value=""BUSCAR"" onclick='VerifyForm" & strFormName & "();' name=""btn" & strFormName & """ id=""btn" & strFormName & """ class=""search_submit"">" & vbCrLf)

                sb.Append("                </form>" & vbCrLf)

                sb.Append("            </div>" & vbCrLf)

                sb.Append("<script language=""javascript"">" & Chr(13))
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

                HttpContext.Current.Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        'Public Function RegistryNotice() As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    Try

        '        sb.Append(" <div class=""registry_notice"">" & vbCrLf)

        '        sb.Append(" <a href=""" & _strProjectPath & "#""><img src=""" & _strProjectPath & "assets/imgs/gral/btn_round_bg_join_bkg.png"" alt=""unete hoy"" /></a>" & vbCrLf)

        '        sb.Append(" </div>" & vbCrLf)

        '        Return sb.ToString()

        '    Catch ex As Exception

        '        HttpContext.Current.Trace.Warn(ex.ToString())

        '        Throw ex

        '    Finally

        '        sb = Nothing

        '    End Try

        'End Function





        Public Function Encuesta() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim objSurvey As New Paralideres.Survey(Date.Today)

            Try

                sb.Append("                 <div id=""DivSurvey""class=""sb_content sv_box clearfix"">" & vbCrLf)

                sb.Append("                    	<!-- Starts Survey Widget Form -->" & vbCrLf)

                sb.Append(objSurvey.DisplayCurrentSurvey2010)

                sb.Append("                        <!-- Ends survey widget form -->" & vbCrLf)

                sb.Append("                    </div>" & vbCrLf)


                Return sb.ToString()

            Catch ex As Exception

                HttpContext.Current.Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

                objSurvey = Nothing

            End Try

        End Function

    End Class

End Namespace