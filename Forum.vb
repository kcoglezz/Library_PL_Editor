Namespace ParaLideres

    Public Class Forum

        Private _intForumId As Integer = 0
        Private _strForumName As String = ""
        Private _strForumDesc As String = ""
        Private _intTotalPosts As Integer = 0
        Private _dateLastPost As Date
        Private _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _strPageUrl As String = _project_path & "forum.aspx" 'HttpContext.Current.Request.ServerVariables("URL")
        Private _strBgColor As String = "white"
        Private _strTrail As String = ""

        Public ReadOnly Property ForumName() As String
            Get

                Return _strForumName

            End Get
        End Property

        Public ReadOnly Property ForumDesc() As String
            Get

                Return _strForumDesc

            End Get
        End Property

        Public ReadOnly Property TotalPosts() As Integer
            Get
                Return _intTotalPosts


            End Get
        End Property

        Public ReadOnly Property LastPost() As Date
            Get
                Return _dateLastPost

            End Get
        End Property

        Public ReadOnly Property Trail() As String
            Get

                Return _strTrail

            End Get
        End Property

        Sub New()

        End Sub

        Sub New(ByVal intForumId As Integer)

            _intForumId = intForumId
            setForumInfo()
            _strTrail = CrumTrail()

        End Sub

        Sub setForumInfo()

            Dim reader As System.Data.SqlClient.SqlDataReader

            reader = ParaLideres.GenericDataHandler.GetRecords("sp_GetAllForumByID " & _intForumId)

            If reader.HasRows Then

                reader.Read()

                _strForumName = reader("Forum")
                _strForumDesc = reader("Description")
                _intTotalPosts = reader("Posts")
                _dateLastPost = reader("LastPost")

            End If

            reader.Close()

        End Sub

        Public Function DisplayForum(ByVal intIndex As Integer) As String

            If _intForumId > 0 Then  'display all the posts for this forum

                Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
                Dim reader As System.Data.SqlClient.SqlDataReader
                Dim intMessageId As Integer
                Dim strTitle As String
                Dim strPostedBy As String
                Dim strPostedOn As String
                Dim strScript As String = ""
                Dim strBgColor As String = "white"
                Dim intTotalReplies As Integer = 0
                Dim strLinkReply As String = "<p align=center><a href=""replyforum.aspx?forum_id=" & _intForumId & "&parent_id=0"">Crea Nuevo Mensaje</a></p>"
                Dim intTotalRecords As Integer = 0

                'Get Total number of posts

                reader = ParaLideres.GenericDataHandler.GetRecords("Select count(MessageId) from Messages where ForumID = " & _intForumId)

                reader.Read()

                intTotalRecords = reader(0)

                'Get Messages Posted for this Forum

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMessagesByForumID " & _intForumId & "," & intIndex)

                If reader.HasRows Then

                    sb.Append(strLinkReply)

                    sb.Append("<p align=center>")

                    'Table Header
                    sb.Append("<table cellpadding=3 cellspacing=0 rules=""columns"" bordercolor=""Black"" border=""0"" style=""border-color:Black;border-width:0px;border-style:solid;width:500px;border-collapse:collapse;"">")

                    sb.Append("<tr bgcolor=#B7D983 class=Estilo3 valign=top align=center><td nowrap>Tema</td><td nowrap>Respuestas</td><td nowrap>Publicado Por</td><td nowrap>Publicado En</td></tr>")

                    Do While reader.Read()

                        'strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intMessageId = reader(0)
                        strPostedBy = reader(1)
                        strTitle = "<a href=message.aspx?message_id=" & intMessageId & "&forum_id=" & _intForumId & ">" & reader(2) & "</a>"
                        strPostedOn = Functions.FormatHispanicDateTime(reader(3))
                        intTotalReplies = CountReplies(intMessageId)

                        sb.Append("<tr class=Estilo1  " & strScript & "><td>" & strTitle & "</td><td>" & intTotalReplies & "</td><td>" & strPostedBy & "</td><td>" & strPostedOn & "</td></tr>")

                        If strBgColor = "white" Then strBgColor = "#D0F29C" Else strBgColor = "white"

                        sb.Append(HorRow(4))

                    Loop

                    sb.Append("</table>")

                    sb.Append("</p>")

                    sb.Append("<p align=center>" & BottomMenu(intTotalRecords, intIndex) & "</p>")

                End If

                reader.Close()


                Return sb.ToString()

            Else 'display main forum list

                Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
                Dim reader As System.Data.SqlClient.SqlDataReader
                Dim intForumId As Integer
                Dim strForumName As String
                Dim strDesc As String
                'Dim strPostedOn As String
                'Dim intTotalPosts As Integer

                Dim strScript As String = ""
                Dim strBgColor As String = "white"

                reader = ParaLideres.GenericDataHandler.GetRecords("sp_GetAllForums")

                If reader.HasRows Then

                    sb.Append("<p align=center>")

                    'Table Header
                    sb.Append("<table cellpadding=3 cellspacing=0 rules=""columns"" bordercolor=""Black"" border=""0"" style=""border-color:Black;border-width:0px;border-style:solid;width:500px;border-collapse:collapse;"">")

                    sb.Append("<tr bgcolor=#B7D983 class=Estilo3><td>Foro</td><td>Descripci&#243;n</td></tr>")

                    Do While reader.Read()

                        'strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & strBgColor & ";"""

                        intForumId = reader(0)
                        strForumName = "<a href=" & _project_path & "forum.aspx?forum_id=" & intForumId & ">" & reader(1) & "</a>"
                        strDesc = Functions.ReplaceCR(reader(2))
                        'intTotalPosts = reader(3)
                        'strPostedOn = FormatHispanicDateTime(reader(4))

                        sb.Append("<tr class=Estilo1  " & strScript & "><td valign=top>" & strForumName & "</td><td valign=top>" & strDesc & "</td></tr>")

                        If strBgColor = "white" Then strBgColor = "#D0F29C" Else strBgColor = "white"

                        sb.Append(HorRow())

                    Loop

                    sb.Append("</table>")

                    sb.Append("</p>")

                End If

                reader.Close()

                Return sb.ToString()


            End If


        End Function

        Public Function BottomMenu(ByVal intRecords As Integer, ByVal intIndex As Integer) As String

            Dim intPages As Integer = 0
            Dim intPageSize As Integer = 10
            Dim intCurrentPage As Integer = 0
            Dim _strPrev As String = ""
            Dim _strNext As String = ""
            Dim strNumericMenu As String = ""
            Dim intX As Integer = 0
            Dim intMenuStart As Integer = 0
            Dim intMenuEnd As Integer = 0

            If intMenuEnd > (intPages - 1) Then intMenuEnd = (intPages - 1)


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

                    strNumericMenu = strNumericMenu & " <a href=" & _strPageUrl & "?index=" & (intX * intPageSize) + 1 & "&forum_id=" & _intForumId & ">" & intX + 1 & "</a> "

                End If

            Next

            If intPages > 1 Then

                Select Case intCurrentPage

                    Case Is <= 1

                        _strPrev = "<< Previo " & strNumericMenu
                        _strNext = "<a href=" & _strPageUrl & "?index=" & intIndex + intPageSize & "&forum_id=" & _intForumId & ">Siguiente >></a>"

                    Case Is >= intPages

                        _strPrev = "<a href=" & _strPageUrl & "?index=" & intIndex - intPageSize & "&forum_id=" & _intForumId & "><< Previo</a> " & strNumericMenu
                        _strNext = "Siguiente >>"

                    Case Else

                        _strPrev = "<a href=" & _strPageUrl & "?index=" & intIndex - intPageSize & "&forum_id=" & _intForumId & "><< Previo</a> " & strNumericMenu
                        _strNext = "<a href=" & _strPageUrl & "?index=" & intIndex + intPageSize & "&forum_id=" & _intForumId & ">Siguiente >></a>"

                End Select

            End If

            Return _strPrev & " " & _strNext

        End Function

        Public Function DisplayThread(ByVal intMessageId As Integer) As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim strPostedBy As String
            Dim strSubject As String
            Dim strBody As String
            Dim intChildren As Integer
            Dim strPosted As String
            Dim strScript As String = ""
            Dim strLinkReply As String = ""


            reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetParentMessageByID " & intMessageId)

            If reader.HasRows Then

                reader.Read()

                sb.Append("<p align=center>")

                'Table Header
                sb.Append("<table cellpadding=3 cellspacing=0 rules=""columns"" bordercolor=""Black"" border=""0"" style=""border-color:Black;border-width:0px;border-style:solid;width:500px;border-collapse:collapse;"">")

                'sb.Append("<tr bgcolor=#B7D983 class=BOLD><td>Tema</td><td>Respuestas</td><td>Publicado Por</td><td>Publicado En</td></tr>")

                'strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & _strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & _strBgColor & ";"""

                intMessageId = reader("messageid")
                strPostedBy = reader("postedby")
                strSubject = reader("subject")
                strBody = Functions.ReplaceCR(reader("body"))
                strPosted = Functions.FormatHispanicDateTime(reader("mDate"))
                intChildren = reader("children")
                _intForumId = reader("forumid")

                setForumInfo()

                _strTrail = CrumTrail(intMessageId, strSubject)

                strLinkReply = "<a href=""replyforum.aspx?forum_id=" & _intForumId & "&parent_id=" & intMessageId & """><img src=" & _project_path & "images/reply.gif border=0> Responde a este mensaje</a>"

                sb.Append(DisplayMessageRow(strSubject, strPostedBy, strPosted, strBody, strScript, strLinkReply, 0))

                If _strBgColor = "white" Then _strBgColor = "#D0F29C" Else _strBgColor = "white"

                If intChildren > 0 Then sb.Append(GetMessages(intMessageId, 1))

                sb.Append("</table>")

                reader.Close()

            End If

            Return sb.ToString()

        End Function

        Public Sub DeleteThread(ByVal intParentId As Integer)

            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intMessageId As Integer = 0
            Dim intChildren As Integer

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMessagesByParentID " & intParentId)

                If reader.HasRows Then

                    Do While reader.Read()

                        intMessageId = reader("messageid")
                        intChildren = reader("children")

                        If intChildren > 0 Then DeleteThread(intMessageId)

                    Loop

                    reader.Close()

                End If

            Catch ex As Exception

                Throw ex

            Finally

                reader = Nothing

            End Try

        End Sub

        Public Function GetMessages(ByVal intParentId As Integer, ByVal intLevel As Integer) As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intMessageId As Integer = 0
            Dim strPostedBy As String
            Dim strSubject As String
            Dim strBody As String
            Dim intChildren As Integer
            Dim strPosted As String
            Dim strScript As String = ""
            Dim strLinkReply As String = ""

            reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetMessagesByParentID " & intParentId)

            If reader.HasRows Then

                sb.Append("<p align=center>")

                Do While reader.Read()

                    'strScript = " onmouseover=""this.style.backgroundColor='lightyellow';this.style.color='red'"" onmouseout=""this.style.backgroundColor='" & _strBgColor & "';this.style.color='black'"" style=""color:Black;background-color:" & _strBgColor & ";"""

                    intMessageId = reader("messageid")
                    strPostedBy = reader("postedby")
                    strSubject = reader("subject")
                    strBody = Functions.ReplaceCR(reader("body"))
                    strPosted = Functions.FormatHispanicDateTime(reader("mDate"))
                    intChildren = reader("children")

                    strLinkReply = "<a href=""replyforum.aspx?forum_id=" & _intForumId & "&parent_id=" & intMessageId & """><img src=" & _project_path & "images/reply.gif border=0>Responde a este mensaje</a>"

                    sb.Append(DisplayMessageRow(strSubject, strPostedBy, strPosted, strBody, strScript, strLinkReply, intLevel))

                    If _strBgColor = "white" Then _strBgColor = "#D0F29C" Else _strBgColor = "white"

                    If intChildren > 0 Then sb.Append(GetMessages(intMessageId, intLevel + 1))

                Loop

                reader.Close()

            End If

            Return sb.ToString()

        End Function

        Private Function DisplayMessageRow(ByVal strSubject As String, ByVal strPostedBy As String, ByVal strPosted As String, ByVal strBody As String, ByVal strScript As String, ByVal strLinkReply As String, ByVal intLevel As Integer) As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")

            Dim intWidth As Integer = 500

            Dim intBorderWidth As Integer = 1

            Dim strBorderColor = "#CCCCCC"

            If intLevel = 0 Then

                strBorderColor = "black"
                intBorderWidth = 2

            End If


            Dim style1 As String = "border-color:" & strBorderColor & ";border-width:" & intBorderWidth & "px;border-style:solid;width:"
            Dim style2 As String = "px;border-collapse:collapse"


            sb.Append("<tr class=Estilo1 " & strScript & ">")
            sb.Append("<td align=left valign=top >")

            sb.Append("<table cellpadding=2 cellspacing=0 rules=""rows"" bordercolor=""Black"" style=" & style1 & intWidth & style2 & ">")

            sb.Append("<tr class=Estilo1 " & strScript & ">")

            sb.Append("<td align=left valign=top >")
            sb.Append("<img src=" & _project_path & "images/spacer.gif width=" & (intLevel * 30) & " height=9 border=0>")
            sb.Append("</td>")

            sb.Append("<td valign=top><b>" & strSubject & "</b>")
            sb.Append("<p><i>Publicado Por:" & strPostedBy & " el " & strPosted & "</i></p>")
            sb.Append("<p>" & strBody)
            sb.Append("<p align=center>" & strLinkReply & "<br><br></p>")
            sb.Append("</td>")

            sb.Append("</tr>")

            sb.Append(HorRow())

            sb.Append("</table>")

            sb.Append("</td>")
            sb.Append("</tr>")

            Return sb.ToString()

        End Function


        Private Function DisplayMessage(ByVal strSubject As String, ByVal strPostedBy As String, ByVal strPosted As String, ByVal strBody As String) As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")

            sb.Append("<p class=Estilo1><b>" & strSubject & "</b>")
            sb.Append("<br><i>Publicado Por:" & strPostedBy & " el " & strPosted & "</i></p>")
            sb.Append("<p class=Estilo1>" & strBody)

            Return sb.ToString()

        End Function

        Private Function CountReplies(ByVal intMessageId As Integer) As Integer

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intTotalChildren As Integer = 0
            Dim intChildId As Integer = 0

            reader = ParaLideres.GenericDataHandler.GetRecords("sp_CountChildrenByMainParentID " & intMessageId)

            If reader.HasRows Then

                Do While reader.Read

                    intChildId = reader(0)

                    intTotalChildren = intTotalChildren + 1

                    intTotalChildren = intTotalChildren + CountReplies(intChildId)

                Loop

                reader.Close()

            End If

            Return intTotalChildren

        End Function

        Public Function CrumTrail(Optional ByVal intParentId As Integer = 0, Optional ByVal strParentSubject As String = "") As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
            Dim str As String

            sb.Append(">><a href=" & _project_path & "forum.aspx>Foros</a>")

            If _intForumId > 0 Then

                sb.Append(">><a href=" & _project_path & "forum.aspx?forum_id=" & _intForumId & ">" & ForumName & "</a>")

            End If

            If intParentId > 0 Then

                sb.Append(">><a href=" & _project_path & "message.aspx?message_id=" & intParentId & ">" & strParentSubject & "</a>")

            End If

            Return sb.ToString()

        End Function


        Public Function DisplayReplyForm(Optional ByVal intParentId As Integer = 0) As String

            Dim objFrm As ParaLideres.FormControls.GenericForm = New ParaLideres.FormControls.GenericForm("frmForum")
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
            Dim objParentMsg As DataLayer.messages

            Dim strSubject As String = ""

            If intParentId > 0 Then

                objParentMsg = New DataLayer.messages(intParentId)

                strSubject = objParentMsg.getSubject

                sb.Append(DisplayMessage(strSubject, objParentMsg.getPostedby, Functions.FormatHispanicDateTime(objParentMsg.getMdate), objParentMsg.getBody))

                strSubject = "RE: " & strSubject

            End If

            Try

                sb.Append(objFrm.FormAction("post_message.aspx"))
                sb.Append(objFrm.FormHidden("forum_id", _intForumId))
                sb.Append(objFrm.FormHidden("in_reply_to", intParentId))
                sb.Append(objFrm.FormTextBox("T&#237;tulo", "strTitle", strSubject, 70, "Ingresa el t&#237;tulo de tu mensaje", True))
                'sb.Append(objFrm.FormTextAreaPlus("Mensaje", "strMessage", "", 450, 300, "Ingresa tu mensaje", "Arial", 3, False))
                sb.Append(objFrm.FormTextArea("Mensaje", "strMessage", "", 15, 70, "Ingresa tu mensaje", True, 4000))
                sb.Append(objFrm.FormEnd("Enviar", "strTitle"))

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                objFrm = Nothing
                sb = Nothing
                objParentMsg = Nothing

            End Try

        End Function

        Public Function PostMessage(ByVal intMessageId As Integer, ByVal intForumID As Integer, ByVal intReplyTo As Integer, ByVal strPostedBy As String, ByVal strEmail As String, ByVal strSubject As String, ByVal strBody As String, ByVal datePosted As Date) As String

            Dim objMsg As DataLayer.Messages
            Dim objParent As DataLayer.Messages

            Try

                objMsg = New DataLayer.Messages(intMessageId)

                If intMessageId > 0 Then

                    objMsg.Update(intMessageId, intForumID, intReplyTo, strPostedBy, strEmail, strSubject, strBody, datePosted, 1)

                Else

                    objMsg.Add(intForumID, intReplyTo, strPostedBy, strEmail, strSubject, strBody, datePosted, 1)

                End If

                If intReplyTo > 0 Then

                    objParent = New DataLayer.Messages(intReplyTo)

                    Try

                        ParaLideres.Functions.SendMail(strEmail, objParent.getEmail, "Respuesta publicada en el foro de Para L&#237;deres", objParent.getPostedby & ", este es un mensaje para informarte que " & strPostedBy & " ha publicado una respuesta a tu mensaje " & objParent.getSubject & " en el foro de Para L&#237;deres.  Visita el foro en www.ParaLideres.org para ver este mensaje.")

                    Catch ex As Exception

                        Throw ex

                    End Try

                End If

            Catch ex As Exception

                Throw ex

            Finally

                objMsg = Nothing

                If Not IsNothing(objParent) Then objParent = Nothing

            End Try

        End Function

        Private Function HorRow(Optional ByVal intColSpan As Integer = 2) As String

            Return "<tr><td height=1 bgcolor=#B7D983 colspan=" & intColSpan & "></td></tr>"

        End Function

    End Class

End Namespace
