Imports System
Imports System.Web
Imports System.Web.Mail
Imports System.Text
Imports System.Web.UI

Namespace ParaLideres

    Public Class PT3

        Inherits Page


#Region "Declarations"

        Private _intContentTableWidth As Integer = 824
        Private _intContentTableHeight As Integer = 768
        Private _intPageFormat As PageFormats = PageFormats.NormalPage

        Private _strPateTitle As String = ""
        Private _strPateContent As String = ""
        Private _strBackgroundColor As String = "#3D9394"
        Private _strProjectPath As String = System.Configuration.ConfigurationSettings.AppSettings("ProjectPath")

        Public Enum PageFormats As Integer

            NormalPage = 0
            ExcelFormat = 1
            WordFormat = 2
            PrintFormat = 3
            AjaxFormat = 4

        End Enum


#End Region


#Region "Properties"

        Public Property ProjectPath() As String

            Get

                Return _strProjectPath

            End Get

            Set(ByVal Value As String)

                _strProjectPath = Value

            End Set

        End Property

        Public WriteOnly Property PageContent() As String

            Set(ByVal Value As String)

                _strPateContent = Value

            End Set

        End Property

        Public WriteOnly Property PageTitle() As String

            Set(ByVal Value As String)

                _strPateTitle = Value

            End Set

        End Property

        Public WriteOnly Property PageFormat() As PageFormats
            Set(ByVal value As PageFormats)

                _intPageFormat = value

            End Set
        End Property
#End Region


#Region "Pagina"



        Public Function NormalPage() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("<!--DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"" -->" & vbCrLf)

                sb.Append("<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf)

                sb.Append("<head>" & vbCrLf)

                sb.Append("<title>" & _strPateTitle & "</title>" & vbCrLf)

                sb.Append("<link rel=""stylesheet"" href=""" & _strProjectPath & "v3/styles/styles.css"" type=""text/css"">" & vbCrLf)

                sb.Append("<script language=""JavaScript"" src=""" & _strProjectPath & "v3/javascript/ajax.js"" type=""text/javascript""></script>" & vbCrLf)

                sb.Append("    <style type=""text/css"">" & vbCrLf)

                sb.Append("        .style1" & vbCrLf)

                sb.Append("        {" & vbCrLf)

                sb.Append("            width: 1024px;" & vbCrLf)

                sb.Append("            height: 768px;" & vbCrLf)

                sb.Append("            border: 1px solid #000000;" & vbCrLf)

                sb.Append("            background-color: " & _strBackgroundColor & ";" & vbCrLf)

                sb.Append("        }" & vbCrLf)

                sb.Append("    </style>" & vbCrLf)

                sb.Append("</head>" & vbCrLf)

                sb.Append("<body>" & vbCrLf)

                sb.Append("<div name=""divPanel"" id=""divPanel"" align=""center"" valign=top class=""panel"">")

                sb.Append("<div align=""center"" id=""divAjaxContent"" name=""divAjaxContent"" class=""AjaxDivContent""></div>")

                sb.Append("</div>")

                sb.Append("<div id=""divInstructions"" name=""divInstructions""  class=""hideError"" style=""FONT-FAMILY:Verdana;FONT-SIZE:9pt;BACKGROUND-COLOR:lightgoldenrodyellow;""></div>")

                sb.Append("    <table align=""left"" cellpadding=""0"" cellspacing=""0"" class=""style1"">" & vbCrLf)

                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""100"" height=""768"">&nbsp;</td>" & vbCrLf)

                sb.Append("            <td width=""" & _intContentTableWidth & """ height=""768"" valign=""top"">" & vbCrLf)

                sb.Append("                " & vbCrLf)

                sb.Append("                <table width=" & _intContentTableWidth & " align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)

                sb.Append("                " & vbCrLf)

                sb.Append("                " & vbCrLf)

                sb.Append("                    <tr><td>" & TopOptions() & "</td></tr>                   " & vbCrLf)

                sb.Append("                    " & vbCrLf)

                sb.Append("                    <tr><td>" & MenuAndSearch() & "</td></tr>                                      " & vbCrLf)

                sb.Append("                    " & vbCrLf)

                sb.Append("                    <tr><td>" & Content() & "</td></tr>                   " & vbCrLf)

                sb.Append("                    " & vbCrLf)

                sb.Append("                </table> " & vbCrLf)

                sb.Append("            " & vbCrLf)

                sb.Append("            </td>" & vbCrLf)

                sb.Append("            <td width=""100"" height=""768"">&nbsp;</td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                sb.Append("    </table>" & vbCrLf)

                sb.Append("</body>" & vbCrLf)

                sb.Append("</html>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function AjaxPage() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("<!--DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"" -->" & vbCrLf)

                sb.Append("<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf)

                sb.Append("<head>" & vbCrLf)

                sb.Append("<title>" & _strPateTitle & "</title>" & vbCrLf)

                sb.Append("<link rel=""stylesheet"" href=""" & _strProjectPath & "v3/styles/styles.css"" type=""text/css"">" & vbCrLf)

                sb.Append("<script language=""JavaScript"" src=""" & _strProjectPath & "v3/javascript/ajax.js"" type=""text/javascript""></script>" & vbCrLf)

                sb.Append("</head>" & vbCrLf)

                sb.Append("<body>" & vbCrLf)


                sb.Append("    <table align=""left"" cellpadding=""0"" cellspacing=""0"" style=""widht:507px;height:206px;border:0px;background-color:transparent;"">" & vbCrLf)

                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""507"" height=""206"">" & _strPateContent & "</td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                sb.Append("    </table>" & vbCrLf)

                sb.Append("</body>" & vbCrLf)

                sb.Append("</html>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function TopOptions() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim intHeight As Integer = 0
            Dim intWidth As Integer = 0
            Dim intLogoHeight As Integer = 259
            Dim intLogoWidth As Integer = 317
            Dim intDivHeight As Integer = 206

            Try

                sb.Append("    <table width=" & _intContentTableWidth & " align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)

                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""" & intLogoWidth & """ height=""" & intLogoHeight & """ rowspan=""2""><img src=""" & _strProjectPath & "v3/images/logo.gif""></td>" & vbCrLf)

                intWidth = _intContentTableWidth - intLogoWidth

                intHeight = intLogoHeight - intDivHeight

                sb.Append("            <td width=""" & intWidth & """ height=""" & intHeight & """ valign=""middle"">" & vbCrLf)

                sb.Append("                <a href=""javascript:"" onclick=""ShowAjaxContent('" & _strProjectPath & "v3/lo_ultimo.aspx'," & intWidth & ", " & intDivHeight & ", this, 'divAjaxContent2', false);"" class=""WHITELINK"">LO ULTIMO</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href=""javascript:"" onclick=""ShowAjaxContent('" & _strProjectPath & "v3/todo.aspx'," & intWidth & ", " & intDivHeight & ", this, 'divAjaxContent2', false);"" class=""WHITELINK"">DESTACADO</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href=""javascript:"" onclick=""ShowAjaxContent('" & _strProjectPath & "v3/logon.aspx'," & intWidth & ", " & intDivHeight & ", this, 'divAjaxContent2', false);"" class=""WHITELINK"">REGISTRATE</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href=""javascript:"" onclick=""ShowAjaxContent('" & _strProjectPath & "v3/todo.aspx'," & intWidth & ", " & intDivHeight & ", this, 'divAjaxContent2', false);"" class=""WHITELINK"">SUGERENCIAS</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href=""javascript:"" onclick=""ShowAjaxContent('" & _strProjectPath & "v3/todo.aspx'," & intWidth & ", " & intDivHeight & ", this, 'divAjaxContent2', false);"" class=""WHITELINK"">ETIQUETAS</a>&nbsp;" & vbCrLf)

                sb.Append("            </td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)


                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""" & intWidth & """ height=""" & intDivHeight & """  style=""background-image: url('" & _strProjectPath & "v3/images/table_background.gif');"">" & vbCrLf)

                'AJAX CONTENT 

                sb.Append("<div align=""center"" id=""divAjaxContent2"" name=""divAjaxContent2"" class=""AjaxDivContent"">")

                'sb.Append(_strLeftSideMenu) TODO

                sb.Append("</div>")

                sb.Append("            </td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                sb.Append("    </table>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Private Function MenuAndSearch() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("    <table  width=" & _intContentTableWidth & " align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)

                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""50%"" height=""40"" bgcolor=""#E9E9E9"">" & vbCrLf)

                sb.Append("                 &nbsp;&nbsp;&nbsp;&nbsp;<a href=""" & _strProjectPath & "v3/section.aspx?section_id=20&index=0"" class=""MENU"">CAPACITACION</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href="""" class=""MENU"">RECURSOS</a>&nbsp;" & vbCrLf)

                sb.Append("                <a href="""" class=""MENU"">CURSOS</a>&nbsp;" & vbCrLf)

                sb.Append("            </td>" & vbCrLf)


                sb.Append("            <td  width=""50%"" bgcolor=""#E9E9E9"">" & vbCrLf)

                'sb.Append(SearchForm())

                sb.Append("                <input type=text size=40>&nbsp;&nbsp;<input type=submit value=""buscar"" class=""frmButton"">&nbsp;&nbsp;")

                sb.Append("            </td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                sb.Append("    </table>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function SearchForm() As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim objForm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("search")

            Try

                sb.Append(objForm.FormAction("search.aspx"))

                sb.Append(objForm.FormTextBox("", "page_title", "", 50, "Ingresa par&#225;metros de b&#250;squeda", True, 4, False, False, False, "proc_SearchResultsForAjax"))

                sb.Append(objForm.FormEnd("Buscar", "search"))

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try




        End Function


        Private Function Content() As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                sb.Append("    <table  width=" & _intContentTableWidth & " align=""left"" cellpadding=""10"" cellspacing=""0"" >" & vbCrLf)

                'TITLE ROW
                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""" & _intContentTableWidth & """ height=""20"" bgcolor=""white"" align=""center"" valign=""middle"" class=""PAGETITLE"">" & vbCrLf)

                sb.Append(_strPateTitle)

                sb.Append("            </td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                'CONTENT ROW
                sb.Append("        <tr>" & vbCrLf)

                sb.Append("            <td width=""" & _intContentTableWidth & """ height=""" & _intContentTableHeight & """ bgcolor=""white""  align=""left"" valign=""top""  class=""PAGECONTENT"">" & vbCrLf)

                sb.Append(_strPateContent)

                sb.Append("            </td>" & vbCrLf)

                sb.Append("        </tr>" & vbCrLf)

                sb.Append("    </table>" & vbCrLf)

                Return sb.ToString()

            Catch ex As Exception

                Trace.Warn(ex.ToString())

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

#End Region

#Region "Page Cycle"

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init


            Trace.IsEnabled = False

            EnableViewState = False


        End Sub

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

            If Session("bg_color") = "" Then

                Session("bg_color") = _strBackgroundColor

            Else

                _strBackgroundColor = Session("bg_color")

            End If

            Select Case _intPageFormat

                Case PageFormats.NormalPage

                    writer.Write(NormalPage())

                Case PageFormats.AjaxFormat

                    writer.Write(AjaxPage())

            End Select

        End Sub



#End Region


#Region "HomePage"





#End Region


#Region "Subs"

        Public Sub ShowError(ByVal excError As Exception)

            Dim objErrHandler As New ErrorHandler
            Dim sb As New StringBuilder("")

            Dim blDebugMode As Boolean = False

            Try

                'Try
                '    blDebugMode = CBool(System.Configuration.ConfigurationManager.AppSettings("IsDebugMode"))
                'Catch ex As Exception
                'End Try


                'If excError.Message <> "Thread was being aborted." Then

                '    If Not blDebugMode Then

                '        Functions.SendMail(_support_account, _develop, "Error en Para Lideres", objErrHandler.ReturnHtmlErrorMessage(excError))

                '    End If

                '    Trace.Write(excError.TargetSite.Name)
                '    Trace.Write(excError.Source)
                '    Trace.Write(excError.Message)
                '    Trace.Write(excError.ToString())

                '    sb.Append("<img src=" & _project_path & "images/error.jpg alt=error align=left HSPACE=5 VSPACE=5>Hubo un error en la p&#225;gina que solicitaste.  Nuestro cuerpo t&#233;cnico ha sido informado al respecto.  Puedes regresar a la p&#225;gina anterior y tratar de nuevo.")

                '    If _objUser.getSecurityLevel = 6 Then

                '        sb.Append("<p><b>Informaci&#243;n para Adminstrador</b></p>")
                '        sb.Append("<p><u>Error Information</u><p>" & excError.TargetSite.Name & "<p>" & excError.Source & "<p>" & excError.Message & "<p>" & excError.ToString())

                '    End If

                '    Me.PageTitle = "Error"
                '    Me.PageContent = sb.ToString()

                'End If

            Catch ex As Exception

                Trace.Warn(ex.ToString())

            Finally

                objErrHandler = Nothing
                sb = Nothing

            End Try

        End Sub

#End Region


#Region "Functions"

        Public Function ColorPicker() As String

            Dim sb As System.Text.StringBuilder

            Dim intX As Integer = 0
            Dim intY As Integer = 0
            Dim intColorIndex As Integer = 28

            Dim strColor As String = ""

            Try

                sb = New System.Text.StringBuilder("")

                sb.Append("<p align=center>")

                sb.Append("<table width=""200"" height=""200"" border=1 cellpadding=""0"" cellspacing=""0"">")

                For intX = 0 To 11

                    sb.Append("<tr>")

                    For intY = 0 To 11

                        Try

                            strColor = ReturnKnownColor(intColorIndex)

                        Catch ex As Exception

                            strColor = "#FFFFFF"

                        End Try

                        sb.Append("<td bgcolor=" & strColor & " width=20 height=20><a href=change.aspx?newcolor=" & strColor & "><img src=" & ProjectPath & "v3/images/square.gif width=20 height=20 border=0></a></td>")

                        Trace.Write(strColor & ": " & intColorIndex)

                        intColorIndex = intColorIndex + 1

                    Next

                    sb.Append("</tr>")

                Next

                sb.Append("</table>")

                sb.Append("</p>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function ReturnKnownColor(ByVal intColorIndex As Integer) As String

            Dim objColor As System.Drawing.ColorConverter

            Dim strColor As String = ""

            Try

                objColor = New System.Drawing.ColorConverter

                Try

                    strColor = objColor.ConvertToString(System.Drawing.Color.FromKnownColor(intColorIndex))

                Catch

                    strColor = "Lavender"

                End Try

                Return strColor

            Catch ex As Exception

                Throw ex

            Finally

                objColor = Nothing

            End Try

        End Function

#End Region

    End Class


End Namespace
