Imports System
Imports System.Web.Mail
Imports System.Text
Imports System.Data
Imports System.Web

Namespace ParaLideres

    Public Class ErrorHandler

#Region "Functions"

        Private Function GetVariables(ByVal inputVarCollection As System.Collections.Specialized.NameValueCollection) As DataTable
            Dim objTable As New DataTable
            objTable.Columns.Add("Name", GetType(String))
            objTable.Columns.Add("Value", GetType(String))

            Dim strItem As String
            Dim objRow As DataRow
            For Each strItem In inputVarCollection
                objRow = objTable.NewRow()
                objRow("Name") = strItem
                objRow("Value") = inputVarCollection(strItem)
                objTable.Rows.Add(objRow)
            Next

            Return objTable
        End Function

        Private Function GetCookieVars() As DataTable
            Dim objTable As New DataTable
            objTable.Columns.Add("Name", GetType(String))
            objTable.Columns.Add("Value", GetType(String))

            Dim strItem As String
            Dim objRow As DataRow
            For Each strItem In HttpContext.Current.Request.Cookies
                objRow = objTable.NewRow()
                objRow("Name") = strItem
                objRow("Value") = HttpContext.Current.Request.Cookies(strItem).Value
                objTable.Rows.Add(objRow)
            Next

            Return objTable
        End Function

        Public Function ReturnHtmlErrorMessage(ByVal ex As Exception) As String

            Dim strMessage As New StringBuilder("")

            Try

                strMessage.Append("URL: " & HttpContext.Current.Request.RawUrl() & "<br>")

                strMessage.Append("IP ADDRESS: <a href=http://www.ip2location.com/free.asp?ipaddresses=" & HttpContext.Current.Request.ServerVariables("REMOTE_ADDR") & " target=new>" & HttpContext.Current.Request.ServerVariables("REMOTE_ADDR") & "</a><br>")

                strMessage.Append("Error: " & ex.Message & "<br>")

                If Not IsNothing(ex.TargetSite) Then

                    strMessage.Append("Method: " & ex.TargetSite.Name & "<br>")

                End If

                strMessage.Append("Error Date/Time: " & Date.Now & "<br>")

                Dim objErrInfoDG As ParaLideres.SimpleDataGrid = New ParaLideres.SimpleDataGrid

                strMessage.Append("<TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0><TR><TD STYLE=""background-color: black; color: white;font-weight: bold;font-family: Arial;"">Error information</TD></TR></TABLE>")

                Dim objTable As New DataTable
                objTable.Columns.Add("Name", GetType(String))
                objTable.Columns.Add("Value", GetType(String))

                Dim objRow As DataRow

                objRow = objTable.NewRow()
                objRow("Name") = "Message"
                objRow("Value") = ex.Message
                objTable.Rows.Add(objRow)

                objRow = objTable.NewRow()
                objRow("Name") = "Source"
                objRow("Value") = ex.Source
                objTable.Rows.Add(objRow)

                objRow = objTable.NewRow()
                objRow("Name") = "TargetSite"
                objRow("Value") = ex.TargetSite.ToString()
                objTable.Rows.Add(objRow)

                objRow = objTable.NewRow()
                objRow("Name") = "StackTrace"
                objRow("Value") = ex.StackTrace
                objTable.Rows.Add(objRow)

                strMessage.Append(objErrInfoDG.RenderToStringSimple(objTable))


                strMessage.Append("<BR><BR><TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0><TR><TD COLSPAN=2 STYLE=""background-color: black; color: white;font-weight: bold;font-family: Arial;"">Querystring Collection</TD></TR></TABLE>")

                Dim objQSDG As SimpleDataGrid = New SimpleDataGrid

                'objQSDG.DataSource = GetVariables(HttpContext.Current.Request.QueryString)
                'objQSDG.DataBind()

                strMessage.Append(objQSDG.RenderToStringSimple(GetVariables(HttpContext.Current.Request.QueryString)))

                strMessage.Append("<BR><BR><TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0><TR><TD COLSPAN=2 STYLE=""background-color: black; color: white;font-weight: bold;font-family: Arial;"">Form Collection</TD></TR></TABLE>")

                Dim objFormDG As SimpleDataGrid = New SimpleDataGrid


                'objFormDG.DataSource = GetVariables(HttpContext.Current.Request.Form)
                'objFormDG.DataBind()
                'strMessage.Append(objFormDG.RenderToString)

                strMessage.Append(objFormDG.RenderToStringSimple(GetVariables(HttpContext.Current.Request.Form)))

                strMessage.Append("<BR><BR><TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0><TR><TD COLSPAN=2 STYLE=""background-color: black; color: white;font-weight: bold;font-family: Arial;"">Cookies Collection</TD></TR></TABLE>")

                Dim objCookiesDG As SimpleDataGrid = New SimpleDataGrid

                'objCookiesDG.DataSource = GetCookieVars()
                'objCookiesDG.DataBind()
                strMessage.Append(objCookiesDG.RenderToStringSimple(GetCookieVars()))


                strMessage.Append("<BR><BR><TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0><TR><TD COLSPAN=2 STYLE=""background-color: black; color: white;font-weight: bold;font-family: Arial;"">Server Variables</TD></TR></TABLE>")

                Dim objServVarDG As SimpleDataGrid = New SimpleDataGrid

                'objServVarDG.DataSource = GetVariables(HttpContext.Current.Request.ServerVariables)
                'objServVarDG.DataBind()
                strMessage.Append(objServVarDG.RenderToStringSimple(GetVariables(HttpContext.Current.Request.ServerVariables)))

                Return strMessage.ToString()

            Catch ex1 As Exception


            Finally

                strMessage = Nothing

            End Try

        End Function

#End Region

    End Class

End Namespace
