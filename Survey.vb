Imports System.Data
Imports System.Data.SqlClient
Imports System.Web

Namespace ParaLideres

    Public Class Survey

        Private _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _objsurvey As DataLayer.questions
        Private _arrQuestions(10) As String
        Private _arrAnswers(10) As Integer
        Private _intTotalAnswers As Integer = 0
        Private _intTableWidth As Integer = 150
        Private _intLeftCellWidth As Integer = 110
        Private _intRightCellWidth As Integer = 40

        Public Sub New(ByVal intSurveyId As Integer)

            Try

                _objsurvey = New DataLayer.questions(intSurveyId)

            Catch ex As Exception

                Functions.ShowError(ex)
                Functions.ShowError(ex.InnerException)

            End Try

        End Sub

        Public Sub New(ByVal dateCurrent As Date)

            Dim intSurveyId As Integer = 0

            Dim reader As System.Data.SqlClient.SqlDataReader

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSurveyIdByDate '" & dateCurrent & "'")

                If reader.HasRows Then

                    reader.Read()

                    If Not reader.IsDBNull(0) Then intSurveyId = reader(0)

                End If

                _objsurvey = New DataLayer.questions(intSurveyId)

            Catch ex As Exception

                intSurveyId = 0

                Functions.ShowError(ex)

            Finally

                reader.Close()
                reader = Nothing

            End Try

        End Sub

        Public ReadOnly Property QuestionDesc() As String
            Get
                Return _objsurvey.getQuestionDesc
            End Get
        End Property

        Public WriteOnly Property TableWidth() As Integer
            Set(ByVal Value As Integer)
                _intTableWidth = Value
            End Set
        End Property

        Public WriteOnly Property LeftCellWidth() As Integer
            Set(ByVal Value As Integer)
                _intLeftCellWidth = Value
            End Set
        End Property

        Public WriteOnly Property RightCellWidth() As Integer
            Set(ByVal Value As Integer)
                _intRightCellWidth = Value
            End Set
        End Property

        Private Sub LoadResults(Optional ByVal blIncludeAnswers As Boolean = False)

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intIndex As Integer = 0
            Dim strSql As String = ""

            Try

                _arrQuestions(0) = _objsurvey.getQuestion1
                _arrQuestions(1) = _objsurvey.getQuestion2
                _arrQuestions(2) = _objsurvey.getQuestion3
                _arrQuestions(3) = _objsurvey.getQuestion4
                _arrQuestions(4) = _objsurvey.getQuestion5
                _arrQuestions(5) = _objsurvey.getQuestion6
                _arrQuestions(6) = _objsurvey.getQuestion7
                _arrQuestions(7) = _objsurvey.getQuestion8
                _arrQuestions(8) = _objsurvey.getQuestion9
                _arrQuestions(9) = _objsurvey.getQuestion10

                If blIncludeAnswers Then

                    strSql = "proc_GetSurveyResults " & _objsurvey.getQuestionId
                    'strSql = "sp_GetTotalAnswersByQuestionID " & _objsurvey.getQuestionId

                    reader = ParaLideres.GenericDataHandler.GetRecords(strSql)

                    If reader.HasRows Then

                        reader.Read()

                        For intIndex = 0 To 9

                            If Not reader.IsDBNull(intIndex) Then

                                _arrAnswers(intIndex) = reader(intIndex)

                                HttpContext.Current.Trace.Warn("reader(" & intIndex & "): " & reader(intIndex))

                            Else

                                _arrAnswers(intIndex) = 0

                            End If

                            _intTotalAnswers = _intTotalAnswers + _arrAnswers(intIndex)

                        Next

                    End If

                End If 'If blIncludeAnswers Then

            Catch ex As Exception

                Throw ex

            Finally

                If Not IsNothing(reader) Then

                    reader.Close()
                    reader = Nothing

                End If

            End Try

        End Sub


        Public Function DisplayCurrentSurvey() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim intIndex As Integer = 0
            Dim strFrmName As String = "frmSurvey"
            Dim strFldName As String = "optvote"

            Try

                LoadResults(False)

                If _objsurvey.getQuestionId > 0 Then

                    sb.Append("<form action=post_survey.aspx method=post name=" & strFrmName & " id=" & strFrmName & ">")

                    sb.Append("<input type=hidden name=question_id value=""" & _objsurvey.getQuestionId & """>")

                    sb.Append("<table width=""140"" border=""0"" cellspacing=""0"" cellpadding=""0"">")

                    'sb.Append("<tr>")
                    'sb.Append("<td valign=top><img src=""" & _project_path & "_images/que_piensas.gif"" width=""146"" height=""15""></td>")
                    'sb.Append("</tr>")

                    sb.Append("<tr>")
                    sb.Append("<td height=""37"" colspan=""2""><span class=""Estilo1"">" & _objsurvey.getQuestionDesc & "</span><br><br></td>")
                    sb.Append("</tr>")

                    For intIndex = LBound(_arrQuestions) To UBound(_arrQuestions)

                        If Len(_arrQuestions(intIndex)) > 0 Then

                            sb.Append("<tr>")
                            sb.Append("<td width=""22"" align=""left"" valign=""top""><input type='radio' name='optvote' value='" & intIndex + 1 & "' ></td>")
                            sb.Append("<td width=""118"" align=""left"" valign=""top""><div align=""left""><span class=""Estilo1"">" & _arrQuestions(intIndex) & "</span></div></td>")
                            sb.Append("</tr>")


                            sb.Append("<tr height=1 bgcolor=#B7D983>")
                            sb.Append("<td colspan=2></td>")
                            sb.Append("</tr>")

                        End If

                    Next

                    sb.Append("<tr>")
                    sb.Append("<td colspan=2 align=center><br><input type=button name=btnVote value=""Enviar"" alt=""Enviar""  onclick='VerifyForm" & strFrmName & "();' class=frmbutton  height=25></td>")
                    sb.Append("</tr>")

                    sb.Append("</form>")

                    sb.Append("<script language=javascript>" & Chr(13))
                    sb.Append("<!--" & Chr(13))

                    'function VerifyForm()
                    sb.Append(Chr(13) & "function VerifyForm" & strFrmName & "(){" & Chr(13))

                    sb.Append("if (!RadioCheck" & strFrmName & "('" & strFldName & "')){" & Chr(13))
                    sb.Append("	alert('Por favor selecciona una de las respuestas');" & Chr(13))
                    sb.Append("}" & Chr(13))

                    sb.Append("else {" & Chr(13))

                    sb.Append("document." & strFrmName & ".btnVote.value = 'Espera...';" & Chr(13))
                    sb.Append("document." & strFrmName & ".btnVote.disabled = true;" & Chr(13))
                    sb.Append("document." & strFrmName & ".submit();" & Chr(13))
                    sb.Append("}" & Chr(13)) 'end of second else statement

                    'end verify form function
                    sb.Append("}" & Chr(13))


                    'validate radio buttons
                    sb.Append("function RadioCheck" & strFrmName & "(ps_fld){" & Chr(13))
                    sb.Append("var ischecked = false;" & Chr(13))
                    sb.Append("var num_of_items = 1;" & Chr(13))

                    sb.Append("num_of_items = eval('document." & strFrmName & ".' + ps_fld + '.length');" & Chr(13))

                    sb.Append("	if (isNaN(num_of_items)) {" & Chr(13))
                    sb.Append("		num_of_items = 1;" & Chr(13))
                    sb.Append("	}" & Chr(13))


                    'sb.Append("alert('num of items:' + num_of_items);")

                    'if num = 1
                    sb.Append("	if (num_of_items == 1) {" & Chr(13))

                    sb.Append("	if (eval('document." & strFrmName & ".' + ps_fld + '.checked') == true){" & Chr(13))
                    sb.Append("		ischecked = true;" & Chr(13))
                    sb.Append("	}" & Chr(13))

                    sb.Append("	}" & Chr(13))

                    ' else
                    sb.Append("else	{" & Chr(13))

                    sb.Append("for (i=0; i < num_of_items; i++){" & Chr(13))

                    sb.Append("	if (eval('document." & strFrmName & ".' + ps_fld + '[i].checked') == true){" & Chr(13))
                    sb.Append("		ischecked = true;" & Chr(13))
                    sb.Append("	}" & Chr(13))

                    sb.Append("}" & Chr(13))

                    sb.Append("}" & Chr(13)) 'end else

                    sb.Append("return ischecked;" & Chr(13))

                    sb.Append("}" & Chr(13))
                    'end RadioCheck



                    sb.Append("//-->" & Chr(13))
                    sb.Append("</script>" & Chr(13))


                    sb.Append("</table>")

                    'sb.Append("<p align=center><a href=""" & Functions.ProjectPath & "other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & """ ")

                    'sb.Append(" onmouseover=""ShowAjaxContent('other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & "&format=print&close=y',530,400,this);""")

                    ''sb.Append(" onmouseout=""ClearError(divAjaxContent);"" ")

                    'sb.Append(">Ver Resultados</a></p>")

                    sb.Append("<p align=center>")

                    sb.Append("<a href=""javascript:""")

                    sb.Append(" onclick=""ShowAjaxContent('other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & "&format=print&close=y',530,400,this);""")

                    sb.Append(">Ver Resultados</a>")

                    sb.Append("</p>")

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function DisplayCurrentSurvey2010() As String

            Dim sb As New System.Text.StringBuilder("")
            Dim intIndex As Integer = 0
            Dim strFrmName As String = "frmSurvey"
            Dim strFldName As String = "optvote"

            Try

                LoadResults(False)

                If _objsurvey.getQuestionId > 0 Then

                    sb.Append("                         <form action=""post_survey_ajax.aspx"" method=post name=" & strFrmName & " id=" & strFrmName & ">")

                    sb.Append("                         <input type=hidden name=question_id value=""" & _objsurvey.getQuestionId & """>")

                    sb.Append("                        <p class=""sv_txt"">" & _objsurvey.getQuestionDesc & "</p>" & vbCrLf)

                    sb.Append("                        <ul>" & vbCrLf)

                    For intIndex = LBound(_arrQuestions) To UBound(_arrQuestions)

                        If Len(_arrQuestions(intIndex)) > 0 Then

                            sb.Append("                        <li><label><input type='radio' name='optvote' value='" & intIndex + 1 & "' />" & _arrQuestions(intIndex) & "</label></li>" & vbCrLf)

                        End If

                    Next

                    sb.Append("                        </ul>" & vbCrLf)

                    sb.Append("                        <p>" & vbCrLf)

                    sb.Append("                        <input class=""link_btn_round_md sv_votar"" type=""button"" name=""sv_submit"" value=""Votar!"" onclick='VerifyForm" & strFrmName & "();' />" & vbCrLf)

                    sb.Append("                        </p>" & vbCrLf)

                    sb.Append("                        </form>" & vbCrLf)


                    sb.Append("<script language=javascript>" & Chr(13))

                    sb.Append("<!--" & Chr(13))

                    sb.Append("var intValChecked = 0;" & Chr(13))

                    'function VerifyForm()
                    sb.Append(Chr(13) & "function VerifyForm" & strFrmName & "(){" & Chr(13))

                    sb.Append("if (!RadioCheck" & strFrmName & "('" & strFldName & "')){" & Chr(13))
                    sb.Append("	alert('Por favor selecciona una de las respuestas');" & Chr(13))
                    sb.Append("}" & Chr(13))

                    sb.Append("else {" & Chr(13))

                    sb.Append("document." & strFrmName & ".sv_submit.value = 'Espera...';" & Chr(13))
                    sb.Append("document." & strFrmName & ".sv_submit.disabled = true;" & Chr(13))
                    'sb.Append("document." & strFrmName & ".submit();" & Chr(13))

                    sb.Append("intValChecked = intValChecked + 1;" & Chr(13)) 'add 1 to intValChecked because is not a zero based index

                    sb.Append("ShowAjaxByDiv('DivSurvey',' post_survey_ajax.aspx?question_id=" & _objsurvey.getQuestionId & "&optvote=' + intValChecked);" & Chr(13))

                    sb.Append("}" & Chr(13)) 'end of second else statement

                    'end verify form function
                    sb.Append("}" & Chr(13))


                    'validate radio buttons
                    sb.Append("function RadioCheck" & strFrmName & "(ps_fld){" & Chr(13))
                    sb.Append("var ischecked = false;" & Chr(13))
                    sb.Append("var num_of_items = 1;" & Chr(13))

                    sb.Append("num_of_items = eval('document." & strFrmName & ".' + ps_fld + '.length');" & Chr(13))

                    sb.Append("	if (isNaN(num_of_items)) {" & Chr(13))
                    sb.Append("		num_of_items = 1;" & Chr(13))
                    sb.Append("	}" & Chr(13))


                    'sb.Append("alert('num of items:' + num_of_items);")

                    'if num = 1
                    sb.Append("	if (num_of_items == 1) {" & Chr(13))

                    sb.Append("	if (eval('document." & strFrmName & ".' + ps_fld + '.checked') == true){" & Chr(13))
                    sb.Append("		ischecked = true;" & Chr(13))
                    sb.Append("     intValChecked = 0;" & Chr(13))
                    sb.Append("	}" & Chr(13))

                    sb.Append("	}" & Chr(13))

                    ' else
                    sb.Append("else	{" & Chr(13))

                    sb.Append("for (i=0; i < num_of_items; i++){" & Chr(13))

                    sb.Append("	if (eval('document." & strFrmName & ".' + ps_fld + '[i].checked') == true){" & Chr(13))
                    sb.Append("		ischecked = true;" & Chr(13))
                    sb.Append("     intValChecked = i;" & Chr(13))
                    sb.Append("	}" & Chr(13))

                    sb.Append("}" & Chr(13))

                    sb.Append("}" & Chr(13)) 'end else

                    sb.Append("return ischecked;" & Chr(13))

                    sb.Append("}" & Chr(13))
                    'end RadioCheck

                    sb.Append("//-->" & Chr(13))
                    sb.Append("</script>" & Chr(13))

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function



        Public Function DisplayCurrentQuestion() As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                sb.Append("          <table width=""" & _intTableWidth & """ border=""0"" cellspacing=""0"" cellpadding=""0"">")

                sb.Append("            <tr>")
                sb.Append("              <td height=""37"" colspan=""2""><span class=""Estilo1""><b>" & _objsurvey.getQuestionDesc & "</b></span></td>")
                sb.Append("            </tr>")

                sb.Append("            <tr>")
                sb.Append("              <td height=""37"" colspan=""2"" align=center>")

                'sb.Append("<span class=""Estilo1""><a href=" & ParaLideres.Functions.ProjectPath & "other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & ">Ver Resultados</a>")

                'sb.Append("<a href=""" & Functions.ProjectPath & "other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & """ ")

                sb.Append("<a href=""javascript:""")

                sb.Append(" onclick=""ShowAjaxContent('other_surveys.aspx?survey_id=" & _objsurvey.getQuestionId & "&format=print&close=y',530,400,this);""")

                sb.Append(">Ver Resultados</a>")

                sb.Append("</td>")
                sb.Append("            </tr>")

                sb.Append("            </table>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function

        Public Function DisplayCurrentSurveyResults() As String

            HttpContext.Current.Trace.Write("Init Current Survey: " & Date.Now())

            Dim sb As New System.Text.StringBuilder("")

            Dim intIndex As Integer = 0
            Dim intBarWidth As Integer = 0
            Dim intBarHeight As Integer = 5

            Dim dblPercentage As Double = 0
            Dim dblProportion As Double = 0.5

            Dim blShowHeaderRow As Boolean = True

            Try

                LoadResults(True)

                If Me._intLeftCellWidth > 110 Then

                    dblProportion = 2
                    intBarHeight = 20
                    blShowHeaderRow = False

                End If

                sb.Append("          <table width=""" & _intTableWidth & """ border=""0"" cellspacing=""0"" cellpadding=""0"">")

                If blShowHeaderRow Then

                    sb.Append("            <tr>")
                    sb.Append("              <td height=""37"" colspan=""2""><span class=""Estilo1""><b>" & _objsurvey.getQuestionDesc & "</b></span><br><br></td>")
                    sb.Append("            </tr>")

                Else

                    sb.Append("            <tr>")
                    sb.Append("              <td height=""37"" colspan=""4"" align=center><span class=""Estilo1""><b>Desde: " & Functions.FormatHispanicDateTime(_objsurvey.getDateStart) & " - Hasta: " & Functions.FormatHispanicDateTime(_objsurvey.getDateEnd) & "</b></span><br><br></td>")
                    sb.Append("            </tr>")

                End If


                For intIndex = LBound(_arrQuestions) To UBound(_arrQuestions)

                    If Len(_arrQuestions(intIndex)) > 0 Then

                        If _intTotalAnswers > 0 Then

                            dblPercentage = _arrAnswers(intIndex) / _intTotalAnswers * 100

                            intBarWidth = Int(dblPercentage * dblProportion)

                        End If

                        sb.Append("            <tr>")
                        sb.Append("              <td width=""" & _intLeftCellWidth & """><span class=""Estilo1"">" & _arrQuestions(intIndex) & "</span></td>")
                        sb.Append("              <td width=""" & _intLeftCellWidth & """><img src=" & _project_path & "_images/survey_bar.gif height=" & intBarHeight & " width=" & intBarWidth & " border=1></td>")
                        sb.Append("              <td width=""" & _intRightCellWidth & """  valign=bottom  align=right class=""Estilo1"">" & FormatNumber(_arrAnswers(intIndex), 0) & " votos</td>")
                        sb.Append("              <td width=""" & _intRightCellWidth & """  valign=bottom  align=right class=""Estilo1""><b>" & FormatNumber(dblPercentage, 2) & "%</b></td>")
                        sb.Append("            </tr>")

                        sb.Append("            <tr><td colspan=2 height=3></td></tr>")

                    End If

                Next

                HttpContext.Current.Trace.Write("intTotalVotes: " & _intTotalAnswers)

                'Show Total Number of Votes
                If _intTotalAnswers > 0 Then

                    sb.Append("            <tr>")
                    sb.Append("              <td colspan=5 class=""Estilo1"" align=center><b>N&#250;mero total de votos: " & _intTotalAnswers & "</b></td>")
                    sb.Append("            </tr>")

                End If

                HttpContext.Current.Trace.Write("after total votos")

                sb.Append("          </table>")

                Return sb.ToString()

                HttpContext.Current.Trace.Write("End Current Survey: " & Date.Now())

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function DisplayCurrentSurveyResultsAjax() As String

            HttpContext.Current.Trace.Write("Init Current Survey: " & Date.Now())

            Dim sb As New System.Text.StringBuilder("")

            Dim intIndex As Integer = 0
            Dim intBarWidth As Integer = 0
            Dim intBarHeight As Integer = 5

            Dim dblPercentage As Double = 0
            Dim dblProportion As Double = 0.5

            Dim blShowHeaderRow As Boolean = True

            Try

                LoadResults(True)

                If Me._intLeftCellWidth > 110 Then

                    dblProportion = 2
                    intBarHeight = 20
                    blShowHeaderRow = False

                End If

                sb.Append("          <table width=""210"" border=""0"" cellspacing=""0"" cellpadding=""0"">")

                'If blShowHeaderRow Then

                sb.Append("            <tr>")
                sb.Append("              <td colspan=""2""><b>" & _objsurvey.getQuestionDesc & "</b></td>")
                sb.Append("            </tr>")

                'Else

                '    sb.Append("            <tr>")
                '    sb.Append("              <td colspan=""3""><b>Desde: " & Functions.FormatHispanicDateTime(_objsurvey.getDateStart) & " - Hasta: " & Functions.FormatHispanicDateTime(_objsurvey.getDateEnd) & "</b></td>")
                '    sb.Append("            </tr>")

                'End If


                For intIndex = LBound(_arrQuestions) To UBound(_arrQuestions)

                    If Len(_arrQuestions(intIndex)) > 0 Then

                        If _intTotalAnswers > 0 Then

                            dblPercentage = _arrAnswers(intIndex) / _intTotalAnswers * 100

                            intBarWidth = Int(dblPercentage * dblProportion)

                        End If

                        sb.Append("            <tr>")
                        sb.Append("              <td valign=""bottom""  align=""left"">" & _arrQuestions(intIndex) & "</td>") 'width=""" & _intLeftCellWidth & """
                        'sb.Append("              <td width=""" & _intLeftCellWidth & """><img src=" & _project_path & "_images/survey_bar.gif height=" & intBarHeight & " width=" & intBarWidth & " border=1></td>")
                        'sb.Append("              <td valign=""bottom""  align=""right"">" & FormatNumber(_arrAnswers(intIndex), 0) & " votos</td>")
                        sb.Append("              <td valign=""bottom""  align=""right""><b>" & FormatNumber(dblPercentage, 2) & "%</b></td>")
                        sb.Append("            </tr>")

                    End If

                Next

                HttpContext.Current.Trace.Write("intTotalVotes: " & _intTotalAnswers)

                'Show Total Number of Votes
                If _intTotalAnswers > 0 Then

                    sb.Append("            <tr>")
                    sb.Append("              <td colspan=""2"" align=""center""><b>N&#250;mero total de votos: " & _intTotalAnswers & "</b></td>")
                    sb.Append("            </tr>")

                End If

                HttpContext.Current.Trace.Write("after total votos")

                sb.Append("          </table>")

                Return sb.ToString()

                HttpContext.Current.Trace.Write("End Current Survey: " & Date.Now())

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function ShowLink() As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                sb.Append("<p align=""center""><a href=""" & Functions.ProjectPath & "other_surveys.aspx"">Resultados Anteriores</a></p>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

    End Class

End Namespace
