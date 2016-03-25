Imports System.Web.UI
Imports System.Text

Namespace ParaLideres

    Public Class ErrorPageTemplate

        Inherits Page

        Public _support_account As String = System.Web.Configuration.WebConfigurationManager.AppSettings("SupportAccount")
        Public _develop As String = System.Web.Configuration.WebConfigurationManager.AppSettings("DeveloperAccount")


        Private _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _page_content As String
        Private _page_title As String
        Private _strBorderColor As String = "#D09000" 'orange
        Private _strBackGroundColor As String = "#A7D168" 'green
        Private _intWidth As Integer = 530

        Public Property ProjectPath() As String
            Get

                Return _project_path

            End Get

            Set(ByVal Value As String)

                _project_path = Value

            End Set

        End Property

        Public WriteOnly Property PageContent() As String

            Set(ByVal Value As String)

                _page_content = Value

            End Set

        End Property

        Public WriteOnly Property PageTitle() As String

            Set(ByVal Value As String)

                _page_title = Value

            End Set

        End Property


        Public Function PageTemplate() As String 'This is the current version used

            Dim sb As StringBuilder = New StringBuilder("")

            Dim strFlashFile As String = ""

            Try

                sb.Append("<html>")
                sb.Append("<head>")

                sb.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">")

                sb.Append("<meta name=""verify-v1"" content=""lemJ1rsZeV1Q0103quVVmRiXpkvuQIw1Ja6JIF2sSf8="" />")

                sb.Append("<script language=""JavaScript"" src=""" & _project_path & "ajax.js"" type=""text/javascript""></script>" & vbLf)

                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/styles.css"" type=""text/css"">" & vbLf)
                sb.Append("<link rel=""STYLESHEET"" href=""" & _project_path & "Editor/editor.css"" type=""text/css"">" & vbLf)

                sb.Append("<title>" & _page_title & "</title>")

                sb.Append("</head>")

                sb.Append("<body >")

                sb.Append("<table width=""" & _intWidth & """ height=""700"" valign=top align=center cellpadding=""0"" cellspacing=""0""  style=""border-color:" & _strBorderColor & ";border-width:1px;border-style:solid;width:" & _intWidth & "px;"">")

                sb.Append("      <tr>")
                sb.Append("        <td height=""67""><a href=" & _project_path & "home.aspx><img src=""" & _project_path & "_images/header_contenido.jpg"" border=0></a></td>")
                sb.Append("      </tr>")

                sb.Append("      <tr>")
                sb.Append("        <td height=""6""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""10""></td>")
                sb.Append("      </tr>")


                sb.Append("      <tr>")
                sb.Append("        <td height=""3"" align=""left"" background=""" & _project_path & "" & "_images/puntos_horizontales.gif""><img src=""" & _project_path & "_images/puntos_horizontales.gif"" width=""3"" height=""1""></td>")
                sb.Append("      </tr>")


                sb.Append("      <tr>")
                sb.Append("        <td height=""4""><img src=""" & _project_path & "_images/separador_45px.gif"" width=""8"" height=""4""></td>")
                sb.Append("      </tr>")


                sb.Append("      <tr>")
                sb.Append("        <td height=""15""><img src=""" & _project_path & "_images/barra_larga.gif"" width=""530"" height=""15""></td>")
                sb.Append("      </tr>")

                'TITULO
                sb.Append("      <tr>")
                sb.Append("              <td width=""" & _intWidth & """ height=""48"" class=""Estilo4"" align=center valign=top>" & _page_title & "</td>")
                sb.Append("      </tr>")

                'CONTENIDO    
                sb.Append("      <tr>")
                sb.Append("              <td width=""" & _intWidth & """  class=""Estilo1"" align=left valign=top>" & _page_content & "</td>")
                sb.Append("      </tr>")


                sb.Append(" </table>")

                sb.Append("</body>")
                sb.Append("</html>")

                Return sb.ToString()


            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function


        Public Function TranslatePath(ByVal strOriginalPath As String) As String

            Dim strNewUrl As String = ""
            Dim strPassedQuery As String = ""
            Dim strQuery As String = ""
            Dim strPageNotFound As String = ""

            Dim isFixed As Boolean = False

            Dim intFrom As Integer = 0
            Dim intLen As Integer = 0
            Dim intQuestionMark As Integer = 0


            Try


                intQuestionMark = InStrRev(strOriginalPath, "?")
                intFrom = InStr(strOriginalPath, "_")
                intLen = InStrRev(strOriginalPath, ".") - intFrom - 1

                Trace.Write("strOriginalPath:" & strOriginalPath)
                Trace.Write("intFrom:" & intFrom)
                Trace.Write("intLen:" & intLen)
                Trace.Write("intQuestionMark:" & intQuestionMark)

                If intQuestionMark > 0 Then

                    strPassedQuery = Mid(strOriginalPath, intQuestionMark)

                    strOriginalPath = Left(strOriginalPath, intQuestionMark)

                End If


                If intLen > 0 And intFrom > 0 Then

                    strQuery = Mid(strOriginalPath, intFrom + 1, intLen)

                End If

                Trace.Write("strQuery:" & strQuery)
                Trace.Write("strPassedQuery:" & strPassedQuery)
                Trace.Write("strOriginalPath:" & strOriginalPath)

                If InStr(strOriginalPath, "/sections/") > 0 Then

                    strNewUrl = Replace(strOriginalPath, "/sections/", "/")

                    isFixed = True

                    Trace.Write("strNewUrl:" & strNewUrl)

                ElseIf InStr(strOriginalPath, "section") > 0 Then


                    strNewUrl = ProjectPath & "section.aspx?section_id=" & strQuery

                    isFixed = True

                    Trace.Write("strNewUrl:" & strNewUrl)

                ElseIf InStr(strOriginalPath, "pages") > 0 Then


                    strNewUrl = ProjectPath & "article.aspx?page_id=" & strQuery

                    isFixed = True

                    Trace.Write("strNewUrl:" & strNewUrl)


                Else

                    strPageNotFound = Mid(strOriginalPath, InStrRev(strOriginalPath, "/") + 1)

                    Trace.Write("strPageNotFound: " & strPageNotFound)

                    Select Case strPageNotFound

                        Case "home.asp"

                            strNewUrl = ProjectPath & "home.aspx"

                            isFixed = True

                        Case "featured.asp"

                            strNewUrl = ProjectPath & "destacado.aspx"

                            isFixed = True

                        Case Else


                    End Select

                End If

                If strPassedQuery <> "" Then strNewUrl = strNewUrl & "?" & strPassedQuery

                Return strNewUrl

            Catch ex As Exception

                Throw ex

            End Try


        End Function


        Public Sub NotifyAndBlock(ByVal strUrl As String, ByVal strReferrer As String, ByVal intErrorType As Integer)

            Dim objCountryLookup As CountryLookup

            Dim strIPDataPath As String = ParaLideres.Functions.GeoDataFilePath
            Dim strCountryName As String = ""
            Dim strCountryCode As String = ""
            Dim strUserIPAddress As String = ""
            Dim strSubject As String = ""
            Dim strMessage As String = ""

            Try

                objCountryLookup = New CountryLookup(strIPDataPath)

                If Session("IPAddress") = "" Then Session("IPAddress") = Request.UserHostAddress()

                strUserIPAddress = Session("IPAddress")

                strCountryName = objCountryLookup.LookupCountryName(strUserIPAddress)

                strCountryCode = objCountryLookup.LookupCountryCode(strUserIPAddress)

                strMessage = "La P&#225;gina " & strUrl & " no existe.<br>Referrer: " & strReferrer & ".<br>" & strUserIPAddress & ": " & strCountryName ' & " <img src=http://" & Request.ServerVariables("SERVER_NAME") & ProjectPath & "images/flags/" & strCountryCode & ".gif border=0>"


                If strCountryName = "China" Then

                    Try

                        ParaLideres.GenericDataHandler.ExecSQL("proc_InsertBloquedIPAddress '" & Session("IPAddress") & "'")

                        strMessage = strMessage & "<br />Esta direcci&#243;n ha sido bloqueada atomaticamente."

                    Catch ex As Exception

                        strMessage = strMessage & "<p align=center><a href=http://" & Request.ServerVariables("SERVER_NAME") & ProjectPath & "deny_ip_address.aspx?ip_to_block=" & Session("IPAddress") & ">block this IP</a></p>"

                    End Try

                Else

                    strMessage = strMessage & "<p align=center><a href=http://" & Request.ServerVariables("SERVER_NAME") & ProjectPath & "deny_ip_address.aspx?ip_to_block=" & Session("IPAddress") & ">block this IP</a></p>"

                End If


                Select Case intErrorType


                    Case 1 '404 not found

                        strSubject = "P&#225;gina No Encontrada"

                        Me.PageTitle = strSubject
                        Me.PageContent = "La P&#225;gina que buscas no existe y nuestro equipo de apoyo ha sido notificado."

                    Case 2 '403

                        strSubject = "P&#225;gina No Autorizada"


                        Me.PageTitle = strSubject
                        Me.PageContent = "No tienes accesso a la p&#225;gina que buscas, nuestro equipo de apoyo ha sido notificado."

                End Select

                'ParaLideres.Functions.SendMail(_support_account, _develop, strSubject, strMessage)


            Catch ex As Exception

                ParaLideres.Functions.SendMail(_support_account, _develop, "Error BlockUser(str, str, int)", ex.ToString())

            Finally

                objCountryLookup = Nothing

            End Try


        End Sub



        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

            writer.Write(PageTemplate())

        End Sub




        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init

            Dim intFoundBloqued As Integer = 0

            Try

                Try
                    intFoundBloqued = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_GetBloquedIPAddress '" & Session("IPAddress") & "'"))
                Catch ex As Exception
                End Try

                If intFoundBloqued > 0 Then Response.Redirect("http://" & Session("IPAddress"))

            Catch ex As Exception

            End Try

        End Sub
    End Class



End Namespace