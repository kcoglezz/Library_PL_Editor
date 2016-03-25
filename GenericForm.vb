Option Explicit On
Imports System
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web


Namespace ParaLideres.FormControls

    Public Class GenericForm

        Private _form_name As String
        Private sb_validation As StringBuilder = New StringBuilder("")
        Private sb_calendar As StringBuilder = New StringBuilder("")
        Private sb_dhtml As StringBuilder = New StringBuilder("")
        Private sb_onload As StringBuilder = New StringBuilder("")
        Private _today As Date = Today()
        Private _formstrInstructions As String = ""
        Private _strAdditionalValidation As String = ""

        Private _strOnFocusStyle As String = " this.style.backgroundColor='#BCDBEC';this.style.color='Black';"
        Private _strOnBlurStyle As String = " this.style.backgroundColor='Window';this.style.color='WindowText';"

        Private Const _default_date As Date = #1/1/1900#

        Private _xml As System.Xml.XmlNode

        Private _HasDHTMLControls As Boolean = False

        Private mContext As System.Web.HttpContext = System.Web.HttpContext.Current

        Private _project_path As String = System.Configuration.ConfigurationSettings.AppSettings("ProjectPath_editor")

        Public WriteOnly Property FormName() As String
            Set(ByVal Value As String)
                _form_name = Value
            End Set
        End Property

        Public WriteOnly Property FormInstructions() As String
            Set(ByVal Value As String)
                _formstrInstructions = Value
            End Set
        End Property

        Public WriteOnly Property AdditionalScript() As String
            Set(ByVal Value As String)
                _strAdditionalValidation = Value
            End Set
        End Property


        Sub New()

            _form_name = "mform1"

            LoadXml()

        End Sub

        Sub New(ByVal formName As String)

            _form_name = formName

            LoadXml()

        End Sub

        Public Function FormAction(Optional ByVal ps_action As String = "post.aspx", Optional ByVal ps_table_width As Integer = 450, Optional ByVal isMultipart As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder("")

            sb.Append(vbCrLf & "<form action=""" & ps_action & """ method=""post"" id=""" & _form_name & """ name=""" & _form_name & """ ")

            If isMultipart Then sb.Append(" enctype=""multipart/form-data"" ")

            sb.Append(" >" & vbCrLf)

            sb.Append("<table width=" & ps_table_width.ToString & " border=0 >" & vbCrLf)

            'sb.Append("<tr class=frmInst><td></td><td align=left valign=top  nowrap>" & _formstrInstructions & "<br><span class=frmRequired>*</span> Son campos obligatorios<br><br></td></tr>" & vbCrLf)

            'sb.Append("<tr class=frmInst><td></td><td align=left valign=top  nowrap><span class=frmRequired>*</span> Son campos obligatorios</td></tr>" & vbCrLf)

            Return sb.ToString

        End Function

        Public Function FormHidden(ByVal ps_name As String, ByVal ps_value As String) As String

            Dim sb As StringBuilder = New StringBuilder("<input type=hidden name=" & ps_name & " value=""" & ps_value & """>" & vbCrLf)

            Return sb.ToString

        End Function


        Public Function FormFile(ByVal strTitle As String, ByVal strFldName As String, ByVal strFldLength As Integer, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False, Optional ByVal strFileTypes As String = "") As String

            Dim sb As StringBuilder = New StringBuilder

            If strInstructions <> "" Then sb.Append(vbCrLf & "<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")

            If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

            sb.Append(strTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>&nbsp;")
            sb.Append("<input type=""file"" name=""" & strFldName & """ id=""" & strFldName & """ ")

            If strFldLength > 40 Then
                sb.Append("size=45 maxlength=" & strFldLength)
            Else
                sb.Append("size=" & strFldLength & " maxlength=" & strFldLength)
            End If

            If strInstructions <> "" Then

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFldName & ", """ & strInstructions & """, " & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

            End If

            sb.Append(" class=frmInput>")
            sb.Append("</td></tr>" & vbCrLf)

            If blIsRequired Then

                sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If


            If strFileTypes <> "" Then

                Dim arrFileTypes As String() = Split(strFileTypes, ",")

                Dim intX As Integer = 0

                sb_validation.Append(vbCrLf & "//file type/extension validation for " & strTitle & vbCrLf)

                sb_validation.Append("else if ((document." & _form_name & "." & strFldName & ".value != """") && ")

                For intX = 0 To UBound(arrFileTypes)

                    sb_validation.Append("(InStr(document." & _form_name & "." & strFldName & ".value,'" & arrFileTypes(intX) & "') == 0) && ")

                Next

                sb_validation.Remove(sb_validation.Length - 3, 3)

                sb_validation.Append("){" & vbCrLf)

                'sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value != """") {" & vbCrLf)

                'sb_validation.Append("alert(document." & _form_name & "." & strFldName & ".value);" & vbCrLf)

                'sb_validation.Append("alert(InStr(document." & _form_name & "." & strFldName & ".value,'" & arrFileTypes(0) & "'));" & vbCrLf)

                'sb_validation.Append("alert(InStr(document." & _form_name & "." & strFldName & ".value,'" & arrFileTypes(1) & "'));" & vbCrLf)

                sb_validation.Append("alert('El documento debe tener terminar en uno de los siguientes formatos: " & Replace(strFileTypes, ",", " o ") & "');" & vbCrLf)

                sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFldName & ".select();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)




            End If



            Return sb.ToString

        End Function

        Public Function FormTextBox(ByVal strTitle As String, ByVal strFldName As String, ByVal _fldvalue As String, ByVal strFldLength As Integer, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False, Optional ByVal minLength As Integer = 0, Optional ByVal validateNumberic As Boolean = False, Optional ByVal isEmail As Boolean = False, Optional ByVal isCaptchaValidation As Boolean = False, Optional ByVal strAjaxTipsSQL As String = "") As String

            Dim sb As StringBuilder = New StringBuilder

            Dim intImageNumber As Integer
            Dim intMinCharForAjax As Integer = 4

            Try

                If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)


                If isCaptchaValidation Then

                    intImageNumber = Functions.GetRandomNumber(10000, 99999)

                    HttpContext.Current.Session("Captcha") = intImageNumber

                    sb.Append("<tr class=frmInst><td></td><td valign=top><img src=CaptchaImage.aspx?x=" & intImageNumber & " align=top><br>")

                    sb.Append(FormHidden("img", intImageNumber))

                    sb.Append("</td></tr>" & vbCrLf)

                End If

                sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap ")

                If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

                sb.Append(strTitle & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")

                sb.Append("<input type=""text"" name=""" & strFldName & """ id=""" & strFldName & """ value=""" & _fldvalue & """ ")

                If strFldLength > 40 Then

                    sb.Append("size=45 maxlength=" & strFldLength)

                Else

                    sb.Append("size=" & strFldLength & " maxlength=" & strFldLength)

                End If

                If strInstructions <> "" Then

                    strInstructions = FormatString(strInstructions)

                    sb.Append(" onfocus=""ShowInstructions(document." & _form_name & "." & strFldName & ", '" & strInstructions & "'," & strInstructions.Length & ");""")

                    sb.Append(" onblur=""ClearError();""")

                End If

                sb.Append(" onKeyUp=""CheckForKeyReleased('" & strFldName & "','" & strAjaxTipsSQL & "','" & strInstructions & "'," & intMinCharForAjax & ");"" ")


                sb.Append(" class=frmInput>")

                sb.Append("</td></tr>" & vbCrLf)

                If isCaptchaValidation Then

                    sb_validation.Append(vbCrLf & "//Captcha input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value != " & intImageNumber & "){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa el número que se ve en la image de arriba.');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If

                'Makes sure that data has been entered
                If blIsRequired Then

                    sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                    'Validate Minimun Length Data
                    If minLength > 0 Then

                        sb_validation.Append(vbCrLf & "//length input validation for " & strTitle & vbCrLf)
                        sb_validation.Append("else if (GetFieldLength(document." & _form_name & "." & strFldName & ") < " & minLength & "){" & vbCrLf)
                        sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & " que tenga por lo menos " & minLength & " letras.');" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                        sb_validation.Append("}" & vbCrLf)

                    End If 'If minLength > 0 Then

                    'Validate Numeric Data
                    If validateNumberic Then

                        sb_validation.Append(vbCrLf & "//numeric input validation for " & strTitle & vbCrLf)
                        sb_validation.Append("else if (isNaN(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                        sb_validation.Append("alert('Por favor ingresa un valor numérico para " & strTitle & "');" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                        sb_validation.Append("}" & vbCrLf)

                    End If 'If validateNumberic Then

                End If 'If blIsRequired Then

                If isEmail Then

                    sb_validation.Append(vbCrLf & "//email input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (!isEmail(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa una dirección de e-mail válida para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function FormTextBoxXml(ByVal strFldName As String, ByVal _fldvalue As String, ByVal strFldLength As Integer, Optional ByVal blIsRequired As Boolean = False, Optional ByVal minLength As Integer = 0, Optional ByVal validateNumberic As Boolean = False, Optional ByVal isEmail As Boolean = False) As String

            Dim sb As StringBuilder
            Dim strLabel As String = ""
            Dim strInst As String = ""
            Dim arrVals(1) As String

            Try

                arrVals = FieldValuesFromXml(strFldName)

                strLabel = arrVals(0)
                strInst = arrVals(1)


                sb = New StringBuilder("")

                sb.Append("<input type=""text"" name=""" & strFldName & """ id=""" & strFldName & """ value=""" & _fldvalue & """ ")

                If strFldLength > 40 Then

                    sb.Append("size=45 maxlength=" & strFldLength)

                Else

                    sb.Append("size=" & strFldLength & " maxlength=" & strFldLength)

                End If

                If strInst <> "" Then

                    strInst = FormatString(strInst)

                    sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFldName & ", """ & strInst & """," & strInst.Length & ");'")
                    sb.Append(" onblur='ClearError();'")

                End If


                sb.Append(" class=frmInput>")

                ''Makes sure that data has been entered
                If blIsRequired Then

                    sb_validation.Append(vbCrLf & "//input validation for " & strLabel & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa un valor para " & strLabel & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                    'Validate Minimun Length Data
                    If minLength > 0 Then

                        sb_validation.Append(vbCrLf & "//length input validation for " & strLabel & vbCrLf)
                        sb_validation.Append("else if (GetFieldLength(document." & _form_name & "." & strFldName & ") < " & minLength & "){" & vbCrLf)
                        sb_validation.Append("alert('Por favor ingresa un valor para " & strLabel & " que tenga por lo menos " & minLength & " letras.');" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                        sb_validation.Append("}" & vbCrLf)

                    End If 'If minLength > 0 Then

                    'Validate Numeric Data
                    If validateNumberic Then

                        sb_validation.Append(vbCrLf & "//numeric input validation for " & strLabel & vbCrLf)
                        sb_validation.Append("else if (isNaN(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                        sb_validation.Append("alert('Por favor ingresa un valor numérico para " & strLabel & "');" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                        sb_validation.Append("}" & vbCrLf)

                    End If 'If validateNumberic Then

                End If 'If blIsRequired Then

                'If isEmail Then

                '    sb_validation.Append(vbCrLf & "//email input validation for " & strLabel & vbCrLf)
                '    sb_validation.Append("else if (!isEmail(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                '    sb_validation.Append("alert('Por favor ingresa una dirección de e-mail válida para " & strLabel & "');" & vbCrLf)
                '    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                '    sb_validation.Append("}" & vbCrLf)

                'End If

                Return FieldLayout(strLabel, sb.ToString, strInst, blIsRequired)

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function FormEmailVerifyBox(ByVal strFieldLabel As String, ByVal strFieldName As String, ByVal strValue As String, Optional ByVal strInstructions As String = "") As String

            Dim sb As StringBuilder

            Dim strFieldName2 As String = strFieldName & "2"

            Try

                sb = New StringBuilder("")

                If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

                'E-mail
                sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap class=frmRequired>*" & strFieldLabel & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")

                sb.Append("<input type=""text"" name=""" & strFieldName & """ id=""" & strFieldName & """ value=""" & strValue & """ size=45 maxlength=100 ")

                If strInstructions <> "" Then

                    strInstructions = FormatString(strInstructions)

                    sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFieldName & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                    sb.Append(" onblur='ClearError();'")

                End If

                sb.Append(" class=frmInput>")

                sb.Append("</td></tr>" & vbCrLf)


                'E-mail Verify
                sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap class=frmRequired>*Verifica " & strFieldLabel & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")

                sb.Append("<input type=""text"" name=""" & strFieldName2 & """ id=""" & strFieldName2 & """ value=""" & strValue & """ size=45 maxlength=100 ")

                strInstructions = "Inserte nuevament su dirección de E-mail para verificar que no hay errores"

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFieldName2 & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

                sb.Append(" class=frmInput>")

                sb.Append("</td></tr>" & vbCrLf)

                ''Makes sure that data has been entered
                'sb_validation.Append(vbCrLf & "//input validation for " & strFieldLabel & vbCrLf)
                'sb_validation.Append("else if (document." & _form_name & "." & strFieldName & ".value == """"){" & vbCrLf)
                'sb_validation.Append("alert('Por favor ingresa un valor para " & strFieldLabel & "');" & vbCrLf)
                'sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                'sb_validation.Append("}" & vbCrLf)

                'Makes sure is valid e-mail
                sb_validation.Append(vbCrLf & "//email input validation for " & strFieldLabel & vbCrLf)
                sb_validation.Append("else if (!isEmail(document." & _form_name & "." & strFieldName & ".value)){" & vbCrLf)
                sb_validation.Append("alert('Por favor ingresa una dirección de e-mail válida para " & strFieldLabel & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

                'Makes sure is valid e-mail2
                sb_validation.Append(vbCrLf & "//email input validation for " & strFieldName2 & vbCrLf)
                sb_validation.Append("else if (!isEmail(document." & _form_name & "." & strFieldName2 & ".value)){" & vbCrLf)
                sb_validation.Append("alert('Por favor ingresa una dirección de e-mail válida');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFieldName2 & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

                'Makes sure two e-mails are the same
                sb_validation.Append(vbCrLf & "//compares e-mail values entered" & vbCrLf)
                sb_validation.Append("else if (document." & _form_name & "." & strFieldName & ".value != document." & _form_name & "." & strFieldName2 & ".value){" & vbCrLf)
                sb_validation.Append("alert('Las direcciones de e-mail que ingresaste son diferentes.  Asegúrate de que sean iguales.');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function FormPasswordTextBox(ByVal strTitle As String, ByVal strFldName As String, ByVal _fldvalue As String, ByVal strFldLength As Integer, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False, Optional ByVal minLength As Integer = 0, Optional ByVal validateNumberic As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap ")

            If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

            sb.Append(strTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>&nbsp;")
            sb.Append("<input type=""password"" name=""" & strFldName & """ value=""" & _fldvalue & """ ")
            If strFldLength > 45 Then
                sb.Append("size=45 maxlength=" & strFldLength)
            Else
                sb.Append("size=" & strFldLength & " maxlength=" & strFldLength)
            End If

            If strInstructions <> "" Then

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFldName & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

            End If

            sb.Append(" class=frmInput>")
            sb.Append("</td></tr>" & vbCrLf)

            'Makes sure that data has been entered
            If blIsRequired Then

                sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)


                'Validate Minimun Length Data
                If minLength > 0 Then

                    sb_validation.Append(vbCrLf & "//length input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (GetFieldLength(document." & _form_name & "." & strFldName & ") < " & minLength & "){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & " that is at least " & minLength & " characters long.');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If 'If minLength > 0 Then

                'Validate Numeric Data
                If validateNumberic Then

                    sb_validation.Append(vbCrLf & "//numeric input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (isNaN(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa un valor numérico para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If 'If validateNumberic Then

            End If 'If blIsRequired Then

            Return sb.ToString

        End Function

        Public Function FormChangePasswordTextBox(ByVal objRegUser As DataLayer.reg_users, ByVal strFieldLabel As String, ByVal strFieldName As String, ByVal strOldPassword As String, ByVal intFieldLength As Integer, Optional ByVal strInstructions As String = "", Optional ByVal isRequired As Boolean = False, Optional ByVal intMinLength As Integer = 0) As String

            Dim sb As StringBuilder = New StringBuilder

            Try

                '***************************************** Begin New Password Field ********************************************

                If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

                sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap ")

                If isRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

                sb.Append(strFieldLabel & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")
                sb.Append("<input type=""password"" name=""" & strFieldName & """ value="""" ")
                If intFieldLength > 45 Then
                    sb.Append("size=45 maxlength=" & intFieldLength)
                Else
                    sb.Append("size=" & intFieldLength & " maxlength=" & intFieldLength)
                End If

                If strInstructions <> "" Then

                    sb.Append(" onfocus=""ShowInstructions(document." & _form_name & "." & strFieldName & ", '" & strInstructions & "'," & strInstructions.Length & ");" & _strOnFocusStyle & """")
                    sb.Append(" onblur=""ClearError(divInstructions);" & _strOnBlurStyle & """")

                End If

                'if the enter key is pressed then submits the form
                sb.Append(" onKeyPress=""SubmitOnEnter" & _form_name & "();"" ")

                sb.Append(" class=frmInput>")
                sb.Append("</td></tr>" & vbCrLf)

                'separator
                'sb.Append(FormFldSeparator())

                '***************************************** End New Password Field ********************************************


                '***************************************** Begin Confirm New Password Field ********************************************
                sb.Append("<tr class=frmInst><td></td><td>Re-enter your new password</td></tr>" & vbCrLf)


                sb.Append(vbCrLf & "<tr><td align=right valign=top  nowrap ")

                If isRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

                sb.Append("Confirm New Password:</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")
                sb.Append("<input type=""password"" name=""" & strFieldName & "2"" value="""" ")
                If intFieldLength > 45 Then
                    sb.Append("size=45 maxlength=" & intFieldLength)
                Else
                    sb.Append("size=" & intFieldLength & " maxlength=" & intFieldLength)
                End If

                sb.Append(" onfocus=""ShowInstructions(document." & _form_name & "." & strFieldName & "2, 'Re-enter your new password',30);" & _strOnFocusStyle & """")
                sb.Append(" onblur=""ClearError(divInstructions);" & _strOnBlurStyle & """")

                sb.Append(" class=frmInput>")
                sb.Append("</td></tr>" & vbCrLf)

                'separator
                'sb.Append(FormFldSeparator())

                '***************************************** End Confirm New Password Field ********************************************

                'Makes sure that data has been entered
                If isRequired Then

                    sb_validation.Append(vbCrLf & "//input validation for " & strFieldName & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFieldName & ".value == """"){" & vbCrLf)
                    sb_validation.Append("alert('Please enter a value for " & strFieldLabel & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".select();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)


                    'Validate Minimun Length Data
                    If intMinLength > 0 Then

                        sb_validation.Append(vbCrLf & "//length input validation for " & strFieldName & vbCrLf)
                        sb_validation.Append("else if (GetFieldLength(document." & _form_name & "." & strFieldName & ") < " & intMinLength & "){" & vbCrLf)
                        sb_validation.Append("alert('Please enter a value for " & strFieldLabel & " that is at least " & intMinLength & " characters long.');" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                        sb_validation.Append("document." & _form_name & "." & strFieldName & ".select();" & vbCrLf)
                        sb_validation.Append("}" & vbCrLf)

                    End If 'If intMinLength > 0 Then


                    'Validate New Password against Old password
                    sb_validation.Append(vbCrLf & "//new password validation" & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFieldName & ".value == '" & strOldPassword & "'){" & vbCrLf)
                    sb_validation.Append("alert('Your new password cannot be the same as your old password');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".select();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)


                    'Validate that new password has at least one number and at least one character
                    sb_validation.Append(vbCrLf & "//confirm password has at least 1 number and 1 char" & vbCrLf)
                    sb_validation.Append("else if (HasNumberAndChar(document." & _form_name & "." & strFieldName & ".value) == false)" & vbCrLf)
                    sb_validation.Append("{" & vbCrLf)
                    sb_validation.Append(Chr(9) & Chr(9) & "alert('Your password must be a combination of numbers and letters');" & vbCrLf)
                    sb_validation.Append(Chr(9) & Chr(9) & "document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                    sb_validation.Append(Chr(9) & Chr(9) & "document." & _form_name & "." & strFieldName & ".select();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)


                    'Validate New Password excludes name
                    sb_validation.Append(vbCrLf & "//validate new password has no employee name" & vbCrLf)
                    sb_validation.Append("else if( InStr(document." & _form_name & "." & strFieldName & ".value,'" & objRegUser.getFirstName & "') || InStr(document." & _form_name & "." & strFieldName & ".value,'" & objRegUser.getLastName & "') || InStr(document." & _form_name & "." & strFieldName & ".value,'" & Left(objRegUser.getFirstName, 1) & Left(objRegUser.getLastName, 7) & "')){" & vbCrLf)
                    sb_validation.Append("alert('Your new password cannot contain your name');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".focus();" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & ".select();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)


                    'Validate Confirm New Password
                    sb_validation.Append(vbCrLf & "//confirm validation" & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFieldName & "2.value != document." & _form_name & "." & strFieldName & ".value) {" & vbCrLf)
                    sb_validation.Append("alert('The two passwords you enetered do not match');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & "2.focus();" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFieldName & "2.select();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If 'If isRequired Then


                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try


        End Function

        Public Function FormTextArea(ByVal strTitle As String, ByVal strFldName As String, ByVal _fldvalue As String, ByVal _rows As Integer, ByVal _columns As Integer, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False, Optional ByVal max_length As Integer = 4000) As String

            Dim sb As StringBuilder = New StringBuilder

            Try

                If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

                sb.Append("<tr><td valign=top align=right nowrap ")

                If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

                sb.Append(strTitle & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")

                sb.Append("<textarea name=""" & strFldName & """ id=""" & strFldName & """ rows=""" & _rows & """ cols=""" & _columns & """")


                If strInstructions <> "" Then

                    sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFldName & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                    sb.Append(" onblur='ClearError();'")

                End If

                sb.Append(" class=frmInput>" & _fldvalue & "</textarea>")


                sb.Append("</td></tr>" & vbCrLf)

                If blIsRequired Then

                    'validates text area input
                    sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa un valor para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If

                'validates text area maximun length
                sb_validation.Append(vbCrLf & "//length input validation for " & strTitle & vbCrLf)
                sb_validation.Append("else if (GetFieldLength(document." & _form_name & "." & strFldName & ") > " & max_length & "){" & vbCrLf)
                sb_validation.Append("alert('El texto que ingresaste para " & strTitle & " excede el número máximo de letras permitido.  El texto para este campo no puede tener más que " & max_length & " letras.');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function FormSeparator(ByVal strValue As String) As String

            Dim sb As StringBuilder = New StringBuilder("")

            sb.Append("<tr class=frmRequired><td colspan=2>" & strValue & "</td></tr>" & vbCrLf)

            Return sb.ToString

        End Function

        Function FormSingleCheckbox(ByVal strTitle As String, ByVal _name As String, ByVal _value As Integer, Optional ByVal strInstructions As String = "", Optional ByVal is_required As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")

            If is_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

            sb.Append(strTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField >")
            sb.Append(vbCrLf & "<input type=checkbox id=""" & _name & """ name=""" & _name & """ value=""on""")

            If _value = 1 Then sb.Append(" checked ")

            sb.Append(" class=frmInput>" & vbCrLf)

            If is_required Then

                sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                sb_validation.Append("else if (!BoxCheck('" & _name & "')){" & vbCrLf)
                sb_validation.Append("	alert('Por favor marca un valor para " & strTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & _name & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If

            Return sb.ToString

        End Function

        Function FormMultipleCheckbox(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_fldvalue As String, ByVal er_sql As String, Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mDR As SqlDataReader

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")
            If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
            sb.Append(erstrTitle & ":</td>")
            sb.Append("<td align=left valign=top class=frmField>")

            Try
                mDR = ParaLideres.GenericDataHandler.GetRecords(er_sql)

                If Not IsNothing(mDR) Then

                    Do While mDR.Read()

                        sb.Append(vbCrLf & "<input type=checkbox id=""" & er_field & """ name=""" & er_field & """ value=""" & mDR.Item(0) & """")

                        If er_fldvalue <> "" Then

                            If CInt(er_fldvalue) = CInt(mDR.Item(0)) Then sb.Append(" checked ")

                        End If

                        sb.Append(" class=frmInput>" & mDR.Item(1) & "<br><br>" & vbCrLf)

                    Loop

                    mDR.Close()
                Else

                    sb.Append(vbCrLf & "Error trying to retrieve data. Data Reader is empty")


                End If

            Catch e As SqlException

                sb.Append(e.Message)

            Finally

                mDR = Nothing

            End Try

            sb.Append("</td></tr>" & vbCrLf)

            If ps_required Then

                sb_validation.Append("else if (!BoxCheck('" & er_field & "')){" & vbCrLf)
                sb_validation.Append("	alert('Por favor marca un valor para " & erstrTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & "[1].focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If

            Return sb.ToString

        End Function

        Public Function FormMultipleRadio(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_fldvalue As String, ByVal er_sql As String, Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mDR As SqlDataReader

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")
            If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
            sb.Append(erstrTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>")

            Try
                mDR = ParaLideres.GenericDataHandler.GetRecords(er_sql)


                If mDR.HasRows Then


                    Do While mDR.Read()

                        sb.Append(vbCrLf & "<input type=radio id=""" & er_field & """ name=""" & er_field & """ value=""" & mDR.Item(0) & """")

                        If er_fldvalue <> "" Then
                            If CInt(er_fldvalue) = CInt(mDR.Item(0)) Then sb.Append(" checked ")
                        End If
                        sb.Append(" class=frmInput>" & mDR.Item(1) & "<br><br>" & vbCrLf)

                    Loop

                    mDR.Close()

                Else

                    sb.Append(vbCrLf & "Error trying to retrieve data. Data Reader is empty")

                End If

            Catch e As SqlException

                sb.Append(e.Message)

            Finally

                mDR = Nothing

            End Try

            sb.Append("</td></tr>" & vbCrLf)

            If ps_required Then

                sb_validation.Append("else if (!RadioCheck('" & er_field & "')){" & vbCrLf)
                sb_validation.Append("	alert('Por favor marca un valor para " & erstrTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & "[1].focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If

            Return sb.ToString

        End Function

        Public Function FormDate(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_value As Date, Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False, Optional ByVal start_date As Date = _default_date, Optional ByVal end_date As Date = _default_date) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mIndex As Integer
            Dim arrDaysInMonth() As Integer = {0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

            er_value = CDate(er_value)

            If start_date = _default_date Then start_date = _today.AddYears(-10)
            If end_date = _default_date Then end_date = _today.AddYears(10)

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            Try

                sb.Append("<tr><td align=right valign=top  nowrap ")
                If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
                sb.Append(erstrTitle & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")


                'Day

                sb.Append("<select name=""" & er_field & "_day"" size=1 id=" & er_field & "_day class=frmInput>" & vbCrLf)

                If er_value.Year() = 1900 Then

                    sb.Append("<option value=0 selected>n/a</option>" & vbCrLf)

                    For mIndex = 1 To 31

                        sb.Append("<option value=" & mIndex & ">" & mIndex & "</option>" & vbCrLf)

                    Next

                Else

                    sb.Append("<option value=0>n/a</option>" & vbCrLf)

                    For mIndex = 1 To arrDaysInMonth(er_value.Month())

                        sb.Append("<option value=" & mIndex)

                        If mIndex = er_value.Day Then sb.Append(" selected ")

                        sb.Append(">" & mIndex & "</option>" & vbCrLf)

                    Next
                End If

                sb.Append("</select>" & vbCrLf)

                '---------------------------------------------------------------------------------------------------------------------

                'Month
                sb.Append("<select name=""" & er_field & "_month"" size=1 id=" & er_field & "_month class=frmInput>" & vbCrLf)

                If er_value.Year = 1900 Then

                    sb.Append("<option value=0 selected>n/a</option>" & vbCrLf)

                    For mIndex = 1 To 12

                        sb.Append("<option value=" & mIndex & ">" & Functions.SpanishMonthName(mIndex) & "</option>" & vbCrLf)

                    Next

                Else

                    sb.Append("<option value=0>n/a</option>" & vbCrLf)

                    For mIndex = 1 To 12

                        sb.Append("<option value=" & mIndex)

                        If mIndex = er_value.Month Then sb.Append(" selected ")

                        sb.Append(">" & Functions.SpanishMonthName(mIndex) & "</option>" & vbCrLf)

                    Next

                End If

                sb.Append("</select>" & vbCrLf)

                '---------------------------------------------------------------------------------------------------------------------

                'Year
                sb.Append("<select name=""" & er_field & "_year"" size=1 id=" & er_field & "_year class=frmInput>" & vbCrLf)


                If er_value.Year = 1900 Then

                    sb.Append("<option value=1900 selected>n/a</option>" & vbCrLf)

                    For mIndex = (start_date.Year) To (end_date.Year)

                        sb.Append("<option value=" & mIndex & ">" & mIndex & "</option>" & vbCrLf)

                    Next

                Else

                    sb.Append("<option value=1900>n/a</option>" & vbCrLf)

                    If er_value.Year < start_date.Year Then start_date = er_value.AddYears(-10)

                    For mIndex = (start_date.Year) To (end_date.Year)

                        sb.Append("<option value=" & mIndex)

                        If mIndex = er_value.Year Then sb.Append(" selected ")

                        sb.Append(">" & mIndex & "</option>" & vbCrLf)

                    Next

                End If

                sb.Append("</select>" & vbCrLf)

                '---------------------------------------------------------------------------------------------------------------------

                sb.Append("<a href=javascript:Today('" & er_field & "');>Seleccionar La Fecha de Hoy</a>")

                sb.Append("</td></tr>" & vbCrLf)

                If ps_required Then

                    'verifies user has enter a date
                    sb_validation.Append(vbCrLf & "//date input validation for " & erstrTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & er_field & "_year.value == 1900 ){" & vbCrLf)
                    sb_validation.Append("alert('Por favor selecciona una fecha para " & erstrTitle & "');" & vbCrLf)
                    sb_validation.Append("document." & _form_name & "." & er_field & "_month.focus();" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                    'verifies date is valid
                    'sb_validation.Append(vbCrLf & "//date validation for " & erstrTitle & vbCrLf)
                    'sb_validation.Append("else if (!isDate(document." & _form_name & "." & er_field & "_year.value, document." & _form_name & "." & er_field & "_month.value, document." & _form_name & "." & er_field & "_day.value  ) ){" & vbCrLf)
                    'sb_validation.Append("alert('The date you selected for " & erstrTitle & " is not a valid date.  Please select a valid date');" & vbCrLf)
                    'sb_validation.Append("document." & _form_name & "." & er_field & "_day.focus();" & vbCrLf)
                    'sb_validation.Append("}" & vbCrLf)

                    ' Else 'if user enter a date then validate

                    'verifies date is valid
                    'sb_validation.Append(vbCrLf & "//date validation for " & erstrTitle & vbCrLf)
                    'sb_validation.Append("else if ((document." & _form_name & "." & er_field & "_year.value != '1900' ) && (!isDate(document." & _form_name & "." & er_field & "_year.value, document." & _form_name & "." & er_field & "_month.value, document." & _form_name & "." & er_field & "_day.value  ))) {" & vbCrLf)
                    'sb_validation.Append("alert('The date you selected for " & erstrTitle & " is not a valid date.  Please select a valid date');" & vbCrLf)
                    'sb_validation.Append("document." & _form_name & "." & er_field & "_day.focus();" & vbCrLf)
                    'sb_validation.Append("}" & vbCrLf)

                End If

                'verifies date is valid
                sb_validation.Append(vbCrLf & "//date validation for " & erstrTitle & vbCrLf)
                sb_validation.Append("else if (!isDate(document." & _form_name & "." & er_field & "_year.value, document." & _form_name & "." & er_field & "_month.value, document." & _form_name & "." & er_field & "_day.value  ) ){" & vbCrLf)
                sb_validation.Append("alert('La fecha que elegiste para " & erstrTitle & " no es válida.  Por favor selecciona una fecha que sea válida');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & "_day.focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)


                Return sb.ToString

            Catch e As Exception

                Return e.ToString

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function FormSelect(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_fldvalue As Integer, ByVal er_sql As String, Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False, Optional ByVal ps_default As String = "Selecciona Una Opción") As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mDR As SqlDataReader

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")

            If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

            sb.Append(erstrTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>&nbsp;")

            sb.Append("<select name=""" & er_field & """ size=1 id=" & er_field)

            If strInstructions <> "" Then

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & er_field & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

            End If

            sb.Append(" class=frmInput>" & vbCrLf)

            Try

                mDR = ParaLideres.GenericDataHandler.GetRecords(er_sql)

                If mDR.HasRows Then

                    sb.Append(vbCrLf & "<option value=""0"">")
                    sb.Append(ps_default)
                    sb.Append("</option>" & vbCrLf)

                    Do While mDR.Read()

                        sb.Append(vbCrLf & "<option value=""" & mDR.Item(0) & """")

                        'If er_fldvalue <> 0 Then

                        If er_fldvalue = mDR.GetInt32(0) Then sb.Append(" selected ")

                        'End If

                        sb.Append(">" & mDR.GetString(1) & "</option>" & vbCrLf)

                    Loop

                End If

            Catch exc As Exception

                Functions.ShowError(exc)

            Finally

                mDR.Close()

                mDR = Nothing

            End Try

            sb.Append("</select></td></tr>" & vbCrLf)

            If ps_required Then

                sb_validation.Append("else if (document." & _form_name & "." & er_field & ".value == ""0""){" & vbCrLf)
                sb_validation.Append("alert('Por favor selecciona otro valor para: " & erstrTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If

            Return sb.ToString

        End Function

        Public Function FormSelect(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_fldvalue As String, ByVal er_sql As String, Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False, Optional ByVal ps_default As String = "Selecciona Una Opción") As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mDR As SqlDataReader

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")

            If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

            sb.Append(erstrTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>&nbsp;")

            sb.Append("<select name=""" & er_field & """ size=1 id=" & er_field)

            If strInstructions <> "" Then

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & er_field & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

            End If

            sb.Append(" class=frmInput>" & vbCrLf)

            Try

                mDR = ParaLideres.GenericDataHandler.GetRecords(er_sql)

                If mDR.HasRows Then

                    sb.Append(vbCrLf & "<option value=""0"">")
                    sb.Append(ps_default)
                    sb.Append("</option>" & vbCrLf)

                    Do While mDR.Read()

                        sb.Append(vbCrLf & "<option value=""" & mDR(0) & """")

                        If er_fldvalue = mDR(0) Then sb.Append(" selected ")

                        'End If

                        sb.Append(">" & mDR.GetString(1) & "</option>" & vbCrLf)

                    Loop

                End If


            Catch exc As Exception

                sb.Append(vbCrLf & "<option value=0>Error: " & exc.Source & "</option>")
                sb.Append(vbCrLf & "<option value=0>Error: " & exc.Message & "</option>")
                sb.Append(vbCrLf & "<option value=0>Error: " & exc.InnerException.ToString & "</option>")


            Finally

                mDR.Close()
                mDR = Nothing

            End Try

            sb.Append("</select></td></tr>" & vbCrLf)

            If ps_required Then

                sb_validation.Append("else if (document." & _form_name & "." & er_field & ".value == ""0""){" & vbCrLf)
                sb_validation.Append("alert('Por favor selecciona otro valor para: " & erstrTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)

            End If

            Return sb.ToString

        End Function


        Public Function FormSelectArray(ByVal erstrTitle As String, ByVal er_field As String, ByVal er_fldvalue As Object, ByVal er_array As String(), Optional ByVal strInstructions As String = "", Optional ByVal ps_required As Boolean = False, Optional ByVal _compared_numeric As Boolean = True) As String

            Dim sb As StringBuilder = New StringBuilder("")

            If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            sb.Append("<tr><td align=right valign=top  nowrap ")
            If ps_required Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
            sb.Append(erstrTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>&nbsp;")

            sb.Append("<select name=""" & er_field & """ size=1 id=" & er_field)



            If strInstructions <> "" Then

                'strInstructions = mContext.Server.HtmlEncode(strInstructions)

                sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & er_field & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                sb.Append(" onblur='ClearError();'")

            End If

            sb.Append(" class=frmInput>" & vbCrLf)

            Try

                Try


                    'sb.Append(vbCrLf & "<option value=""0"">")
                    'sb.Append(ps_default)
                    'sb.Append("</option>" & vbCrLf)

                    Dim x As Integer = 0

                    If _compared_numeric Then

                        For x = LBound(er_array) To UBound(er_array)

                            sb.Append(vbCrLf & "<option value=""" & x & """")

                            If er_fldvalue <> 0 Then

                                If er_fldvalue = x Then sb.Append(" selected ")

                            End If

                            sb.Append(">" & er_array(x) & "</option>" & vbCrLf)

                        Next

                    Else

                        For x = LBound(er_array) To UBound(er_array)

                            sb.Append(vbCrLf & "<option value=""" & er_array(x) & """")

                            If er_fldvalue = er_array(x) Then sb.Append(" selected ")

                            sb.Append(">" & er_array(x) & "</option>" & vbCrLf)

                        Next

                    End If


                Catch exc As Exception

                    sb.Append(vbCrLf & "<option value=0>Error: " & exc.Source & "</option>")
                    sb.Append(vbCrLf & "<option value=1>Error: " & exc.Message & "</option>")
                    sb.Append(vbCrLf & "<option value=2>Error: " & exc.InnerException.ToString & "</option>")

                End Try


            Catch e As SqlException

                sb.Append(vbCrLf & "<option value=0>Error:" & e.Message & "</option>")

            Finally


            End Try


            sb.Append("</select></td></tr>" & vbCrLf)

            If ps_required Then
                sb_validation.Append("else if (document." & _form_name & "." & er_field & ".value == ""0""){" & vbCrLf)
                sb_validation.Append("alert('Por favor selecciona otro valor para: " & erstrTitle & "');" & vbCrLf)
                sb_validation.Append("document." & _form_name & "." & er_field & ".focus();" & vbCrLf)
                sb_validation.Append("}" & vbCrLf)
            End If

            Return sb.ToString

        End Function

        'Public Function FormHtmlEditor(ByVal strTitle As String, ByVal strFldName As String, ByVal _fldvalue As String, ByVal _height As Integer, ByVal _width As Integer, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False) As String

        '    Dim sb As StringBuilder = New StringBuilder("")

        '    sb.Append("<tr><td align=right valign=top  nowrap ")
        '    If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
        '    sb.Append(strTitle & ":</td>")

        '    sb.Append("<td align=left valign=top class=frmField>&nbsp;")
        '    sb.Append("<textarea id=" & strFldName & " name=" & strFldName & " style=""display:none;"">This is the text...</textarea>")
        '    sb.Append("<h:HtmlArea id=" & strFldName & "EditorHtmlArea onHtmlChanged=""ha_OnHtmlChanged(this, document.all['" & strFldName & "']);"" style=""height:" & _height.ToString & "px;width:" & _width.ToString & "px;"">")
        '    sb.Append(_fldvalue)
        '    sb.Append("</h:HtmlArea>")
        '    sb.Append("</td></tr>")

        '    sb.Append("<tr><td></td><td class=frmInst>" & strInstructions & "</td></tr>" & vbCrLf)

        '    Return sb.ToString()


        'End Function

        Public Function FormEnd(Optional ByVal _button_value As String = "Enviar", Optional ByVal element_name As String = "", Optional ByVal submit_message As String = "", Optional ByVal browser_name As String = "IE") As String

            Dim sb As StringBuilder = New StringBuilder("")

            Try

                'sb.Append("<tr><td align=right valign=top  nowrap></td><td align=left valign=top class=frmLabel><br><br><input type=button name=btn" & _form_name & " id=btn" & _form_name & " value=""" & _button_value & """ onclick=""VerifyForm" & _form_name & "();""  class=frmButton></td></tr>" & vbCrLf)

                'sb.Append("<tr class=frmInst><td></td><td align=left valign=top  nowrap><span class=frmRequired>*</span> Son campos obligatorios</td></tr>" & vbCrLf)

                sb.Append("<tr><td align=center valign=top class=frmLabel colspan=2><span class=frmRequired>*</span> Son campos obligatorios<br><input type=button name=btn" & _form_name & " id=btn" & _form_name & " value=""" & _button_value & """ onclick=""VerifyForm" & _form_name & "();""  class=frmButton></td></tr>" & vbCrLf)

                'sb.Append("<tr class=frmInst><td></td><td align=left valign=top  nowrap><span class=frmRequired>*</span> Son campos obligatorios</td></tr>" & vbCrLf)

                sb.Append("</table></form>" & vbCrLf)

                sb.Append("<script language=javascript>" & vbCrLf)

                sb.Append("<!--" & vbCrLf)

                sb.Append("var buttonclicks = 0;" & vbCrLf)

                sb.Append("var whitespace = ' \t\n\r';" & vbCrLf)

                sb.Append("var daysInMonth = new Array(1,31,29,31,30,31,30,31,31,30,31,30,31); " & vbCrLf)

                ''Set focus on a specific element of the form
                'If element_name <> "" Then
                '    sb.Append("document.getElementById('" & element_name & "').focus();" & vbCrLf)
                'End If

                'function VerifyForm()
                sb.Append(vbCrLf & "function VerifyForm" & _form_name & "(){" & vbCrLf)

                sb.Append("if (document." & _form_name & ".btn" & _form_name & ".value == """"){" & vbCrLf)
                sb.Append("alert('wrong button');" & vbCrLf)
                sb.Append("}" & vbCrLf)

                '===================================
                sb.Append(sb_validation.ToString & vbCrLf)
                '===================================

                '===================================
                'Add additional validation script generated manually
                If _strAdditionalValidation <> "" Then

                    sb.Append(_strAdditionalValidation & vbCrLf)

                End If
                '===================================

                sb.Append("else {" & vbCrLf)
                sb.Append("buttonclicks ++;" & vbCrLf)

                'enable dates fields before submitting the form
                sb.Append(sb_calendar.ToString())

                sb.Append("if (buttonclicks > 1){" & vbCrLf)
                sb.Append("alert('La informacion ha sido enviada.');" & vbCrLf)
                sb.Append("}" & vbCrLf)
                sb.Append("else {" & vbCrLf)

                'give message to user if submit_message is not ""
                If submit_message <> "" Then sb.Append("alert('" & submit_message & "');" & vbCrLf)

                If _HasDHTMLControls Then

                    sb.Append(sb_dhtml.ToString())

                End If

                sb.Append("document." & _form_name & ".submit();" & vbCrLf)
                sb.Append("document." & _form_name & ".btn" & _form_name & ".disabled = true;" & vbCrLf)
                sb.Append("}" & vbCrLf)

                sb.Append("}" & vbCrLf)

                'end verify form function
                sb.Append("}" & vbCrLf)


                'function to get field length
                sb.Append("function GetFieldLength(callingElement) {" & vbCrLf)
                sb.Append("var mLength = 0;" & vbCrLf)
                sb.Append("var mContent = callingElement.value;			  " & vbCrLf)
                sb.Append("mLength = mContent.length;" & vbCrLf)

                sb.Append("return mLength;" & vbCrLf)

                sb.Append("}" & vbCrLf)
                'end function to get field length


                'function Today()
                sb.Append("function Today(ps_field){" & vbCrLf)
                sb.Append("eval('document." & _form_name & ".' + ps_field + '_year.value=" & _today.Year() & "');" & vbCrLf)
                If _today.Month <> 1 Then
                    sb.Append("eval('document." & _form_name & ".' + ps_field + '_month.value=" & _today.Month & "');" & vbCrLf)
                Else
                    sb.Append("eval('document." & _form_name & ".' + ps_field + '_month.value=1');" & vbCrLf)
                End If
                sb.Append("eval('document." & _form_name & ".' + ps_field + '_day.value=" & _today.Day & "');" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end function Today()


                'validate date function isDate(year, month, day)
                sb.Append("function isDate (year, month, day){" & vbCrLf)
                sb.Append("var intYear = parseInt(year);" & vbCrLf)
                sb.Append("var intMonth = parseInt(month);" & vbCrLf)
                sb.Append("var intDay = parseInt(day);" & vbCrLf)
                'catch invalid days, except for February
                sb.Append("if (intDay > daysInMonth[intMonth]) return false; " & vbCrLf)
                sb.Append("if ((intMonth == 2) && (intDay > daysInFebruary(intYear))) return false;" & vbCrLf)
                sb.Append("return true;" & vbCrLf)
                sb.Append("}" & vbCrLf)

                'catch invalid days, except for February
                'sb.Append("if (intDay > daysInMonth[intMonth]) return false; " & vbCrLf)
                'sb.Append("if ((intMonth == 2) && (intDay > daysInFebruary(intYear))) return false;" & vbCrLf)
                'sb.Append("return true;" & vbCrLf)
                'sb.Append("}//end function isDate" & vbCrLf)
                'end isDate


                'validate date function isDate2(value)
                'sb.Append("function isDate2(psDate){" & vbCrLf)
                'sb.Append("var _date = psDate;" & vbCrLf)
                'sb.Append("_date = Date.Parse(psDate);" & vbCrLf)
                'sb.Append("var intYear = parseInt(getYear(_date));" & vbCrLf)
                'sb.Append("var intMonth = parseInt(getMonth(_date));" & vbCrLf)
                'sb.Append("var intDay = parseInt(getDate(_date));" & vbCrLf)
                'sb.Append("alert (intMonth + ' ' + intDay + ' ' + intYear);")
                'sb.Append("alert (_date);")
                'sb.Append("}" & vbCrLf)


                'check how many days are in February depending on year
                sb.Append("function daysInFebruary (year){" & vbCrLf)
                '// February has 29 days in any year evenly divisible by four,
                '// EXCEPT for centurial years which are not also divisible by 400.
                sb.Append(" return (  ((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0) ) ) ? 29 : 28 );" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end daysInFebruary

                'validate radio buttons
                sb.Append("function RadioCheck(ps_fld){" & vbCrLf)
                sb.Append("var ischecked = false;" & vbCrLf)
                sb.Append("var num_of_items = 0;" & vbCrLf)
                sb.Append("num_of_items = eval('document." & _form_name & ".' + ps_fld + '.length');" & vbCrLf)
                sb.Append("for (i=0; i < num_of_items; i++){" & vbCrLf)
                sb.Append("	if (eval('document." & _form_name & ".' + ps_fld + '[i].checked') == true){" & vbCrLf)
                sb.Append("		ischecked = true;" & vbCrLf)
                sb.Append("	}" & vbCrLf)
                sb.Append("}" & vbCrLf)
                sb.Append("return ischecked;" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end RadioCheck

                'validate checkboxes
                sb.Append("function BoxCheck(ps_fld){" & vbCrLf)
                sb.Append("var ischecked = false;" & vbCrLf)
                sb.Append("var num_of_items = 0;" & vbCrLf)
                sb.Append("num_of_items = eval('document." & _form_name & ".' + ps_fld + '.length');" & vbCrLf)
                sb.Append("for (i=0; i < num_of_items; i++){" & vbCrLf)
                sb.Append("	if (eval('document." & _form_name & ".' + ps_fld + '[i].checked') == true){" & vbCrLf)
                sb.Append("		ischecked = true;" & vbCrLf)
                sb.Append("	}" & vbCrLf)
                sb.Append("}" & vbCrLf)
                sb.Append("return ischecked;" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end BoxCheck

                'function isEmpty
                sb.Append("function isEmpty(s){" & vbCrLf)
                sb.Append("   return ((s == null) || (s.length == 0));" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end isEmpty

                'function isWhitespace()
                sb.Append("function isWhitespace (s){" & vbCrLf)
                sb.Append("    var i;" & vbCrLf)
                sb.Append("    if (isEmpty(s)) return true;" & vbCrLf)
                sb.Append("    for (i = 0; i < s.length; i++) {" & vbCrLf)
                sb.Append("        var c = s.charAt(i);" & vbCrLf)
                sb.Append("        if (whitespace.indexOf(c) == -1) return false;" & vbCrLf)
                sb.Append("    }" & vbCrLf)
                sb.Append("    return true;" & vbCrLf)
                sb.Append("}" & vbCrLf)
                'end isWhitespace

                'validate e-mail address

                sb.Append("function isEmail (s){" & vbCrLf)
                sb.Append("if (isWhitespace(s)) return false;" & vbCrLf)
                sb.Append("var i = 1;" & vbCrLf)
                sb.Append("var sLength = s.length;" & vbCrLf)
                sb.Append("while ((i < sLength) && (s.charAt(i) != '@')){" & vbCrLf)
                sb.Append("   i++" & vbCrLf)
                sb.Append("}" & vbCrLf)

                sb.Append("if ((i >= sLength) || (s.charAt(i) != '@')) return false;" & vbCrLf)
                sb.Append("else i += 2;" & vbCrLf)

                sb.Append("while ((i < sLength) && (s.charAt(i) != '.')){" & vbCrLf)
                sb.Append("   i++" & vbCrLf)
                sb.Append("}" & vbCrLf)

                sb.Append("if ((i >= sLength - 1) || (s.charAt(i) != '.')) return false;" & vbCrLf)

                sb.Append("else return true;" & vbCrLf)
                sb.Append("}" & vbCrLf)

                'end isEmail()


                'function to open new window w/calendar
                sb.Append("function OpenCalendar(fieldName, selectedDate){" & vbCrLf)
                sb.Append("var mURL = '" & _project_path & "calendar.aspx?FormFieldName=' + fieldName + '&SelectedDate=' + selectedDate;" & vbCrLf)

                sb.Append("var str = 'height=180,innerHeight=180,width=180,innerWidth=180;'" & vbCrLf)

                sb.Append("if (window.screen) {" & vbCrLf)
                sb.Append("var ah = screen.availHeight - 30;" & vbCrLf)
                sb.Append("var aw = screen.availWidth - 10;" & vbCrLf)
                sb.Append("var xc = (aw - 160) / 2;" & vbCrLf)
                sb.Append("var yc = (ah - 160) / 2;" & vbCrLf)

                sb.Append("str += 'left=' + xc + ', screenX =' + xc + ',top=' + yc + ', screenY=' + yc;" & vbCrLf)

                sb.Append("}" & vbCrLf)

                sb.Append("newWin = window.open(mURL, 'calendar',str);" & vbCrLf)

                sb.Append("newWin.focus();" & vbCrLf)
                sb.Append("}// end OpenCalendar function" & vbCrLf) 'end function

                'function ShowInstructions
                sb.Append("function ShowInstructions(sField, sErrMsg, iLen){" & vbCrLf)

                sb.Append("var itop;" & vbCrLf)
                sb.Append("var ileft;" & vbCrLf)
                sb.Append("var size;" & vbCrLf)
                sb.Append("var sName;" & vbCrLf)
                sb.Append("var oItem;" & vbCrLf)

                sb.Append("iLen = iLen * 6;" & vbCrLf)

                sb.Append("oItem = sField;" & vbCrLf)
                sb.Append("if(typeof(oItem)!=""object"")" & vbCrLf)
                sb.Append("    return;" & vbCrLf)

                sb.Append("itop = 0;" & vbCrLf)
                sb.Append("ileft = 0;" & vbCrLf)

                sb.Append("itop = oItem.offsetHeight;" & vbCrLf)
                sb.Append("while (oItem.tagName != ""BODY"" ){" & vbCrLf)
                sb.Append("        itop  += oItem.offsetTop;" & vbCrLf)
                sb.Append("        ileft += oItem.offsetLeft;" & vbCrLf)
                sb.Append("    oItem = oItem.offsetParent;" & vbCrLf)
                sb.Append("}" & vbCrLf)

                sb.Append("divInstructions.style.left = ileft + 'px';" & vbCrLf)

                sb.Append("divInstructions.style.top  = itop + 'px'; " & vbCrLf)

                sb.Append("divInstructions.style.width  = iLen + 'pt'; " & vbCrLf)

                sb.Append("divInstructions.style.height  = '15pt'; " & vbCrLf)

                sb.Append("divInstructions.style.border = '1px solid black'; " & vbCrLf)

                sb.Append("divInstructions.style.color = 'black'; " & vbCrLf)

                sb.Append("divInstructions.innerHTML = sErrMsg;" & vbCrLf)

                sb.Append("divInstructions.style.visibility = 'visible';" & vbCrLf)

                sb.Append("}" & vbCrLf)

                'function ClearError
                sb.Append("function ClearError(){" & vbCrLf)
                sb.Append("divInstructions.style.visibility = 'hidden';" & vbCrLf)
                sb.Append("}//end Clear Error function" & vbCrLf)

                If _HasDHTMLControls Then

                    sb.Append(DHTMLTextAreaJavaScript())

                End If

                sb.Append("//-->" & vbCrLf)
                sb.Append("</script>" & vbCrLf)

                'Select Case browser_name

                '    Case "IE"

                '        sb.Append("<span align=left valign=middle name=""divInstructions"" id=""divInstructions"" style=""position:absolute;z-Index:10;left:0;top:0;BORDER-RIGHT: black 1pt solid;BORDER-TOP: black 1pt solid;BORDER-LEFT: black 1pt solid;COLOR:black;BORDER-BOTTOM: black 1pt solid;FONT-FAMILY:Verdana;FONT-SIZE: 9pt;BACKGROUND-COLOR:lightgoldenrodyellow;WIDTH:50px;HEIGHT:15px;"" class='hideError'></span>" & vbCrLf)

                '    Case Else

                'sb.Append("<div align=left valign=middle name=""divInstructions"" id=""divInstructions"" style=""position:absolute;z-Index:10;left:0;top:0;BORDER-RIGHT: black 1pt solid;BORDER-TOP: black 1pt solid;BORDER-LEFT: black 1pt solid;COLOR:black;BORDER-BOTTOM: black 1pt solid;FONT-FAMILY:Verdana;FONT-SIZE: 9pt;BACKGROUND-COLOR:lightgoldenrodyellow;WIDTH:50px;HEIGHT:15px"" class='hideError'></div>" & vbCrLf)

                'End Select

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function


        Public Function FormDateCal(ByVal strTitle As String, ByVal strFldName As String, ByVal _fldvalue As String, Optional ByVal strInstructions As String = "", Optional ByVal blIsRequired As Boolean = False) As String

            Dim sb As StringBuilder = New StringBuilder

            Try

                strInstructions = strInstructions & " Presiona imagen para abrir el calendario y seleccionar una fecha."

                If _fldvalue = "1/1/1900" Then _fldvalue = "N/A"

                If strInstructions <> "" Then sb.Append("<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

                sb.Append("<tr><td align=right valign=top  nowrap ")

                If blIsRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")

                sb.Append(strTitle & ":</td>")

                sb.Append("<td align=left valign=top class=frmField>&nbsp;")
                sb.Append("<input type=""text"" name=""" & strFldName & """ id=""" & strFldName & """ value=""" & _fldvalue & """ size=10 maxlength=10 disabled=""true""")

                If strInstructions <> "" Then

                    sb.Append(" onfocus='ShowInstructions(document." & _form_name & "." & strFldName & ", """ & strInstructions & """," & strInstructions.Length & ");'")
                    sb.Append(" onblur='ClearError();'")

                End If

                sb.Append(" class=frmInput>")

                'link to open new window w/calendar
                sb.Append("&nbsp;&nbsp;<a href=javascript:OpenCalendar('" & strFldName & "','" & _fldvalue & "');><img src=" & _project_path & "images/calendar.gif border=0 align=top alt=""Presiona imagen para abrir el calendario.""></a>")
                'sb.Append("&nbsp;&nbsp;<img src=" & _project_path & "images/calendar.gif onclick='calendario(document." & _form_name & "." & strFldName & ")' border=0 align=top alt=""Presiona imagen para abrir el calendario.""></a>")

                'sb.Append("&nbsp;&nbsp;<img src=" & _project_path & "images/calendar.gif onclick='calendario(document." & _form_name & "." & strFldName & ")' border=0 align=top alt=""Presiona imagen para abrir el calendario.""/></a>")
                'sb.Append("&nbsp;&nbsp;<img src=" & _project_path & "images/calendar.gif onclick='calendario(" & _fldvalue & ")' border=0 align=top alt=""Presiona imagen para abrir el calendario.""></a>")



                sb.Append("</td></tr>" & vbCrLf)

                'enable this txt field before submitting
                sb_calendar.Append("document." & _form_name & "." & strFldName & ".disabled=false;" & vbCrLf)

                'Makes sure that data has been entered
                If blIsRequired Then

                    sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == """"){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa una fecha para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                    sb_validation.Append(vbCrLf & "//input validation for " & strTitle & vbCrLf)
                    sb_validation.Append("else if (document." & _form_name & "." & strFldName & ".value == 'N/A'){" & vbCrLf)
                    sb_validation.Append("alert('Por favor ingresa una fecha para " & strTitle & "');" & vbCrLf)
                    sb_validation.Append("}" & vbCrLf)

                End If 'If blIsRequired Then

                'verifies date is valid

                'sb_validation.Append(vbCrLf & "//date validation for " & strTitle & vbCrLf)
                'sb_validation.Append("else if (!isDate2(document." & _form_name & "." & strFldName & ".value)){" & vbCrLf)
                'sb_validation.Append("alert('The date you selected for " & strTitle & " is not a valid date.  Please select a valid date');" & vbCrLf)
                'sb_validation.Append("document." & _form_name & "." & strFldName & ".focus();" & vbCrLf)
                'sb_validation.Append("}" & vbCrLf)

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Public Function FormatString(ByVal ps_string As String) As String

            'ps_string = mContext.Server.HtmlEncode(ps_string)
            ps_string = Replace(ps_string, Chr(34), "&#34;")
            ps_string = Replace(ps_string, Chr(39), "&#39;")

            Return ps_string

        End Function

        Public Function FormTextAreaPlus(ByVal strTitle As String, ByVal strFldName As String, ByVal strFldVal As String, ByVal intWidth As Integer, ByVal intHeight As Integer, Optional ByVal strInstructions As String = "", Optional ByVal strFonts As String = "Arial,Verdana,Roman", Optional ByVal intFontMaxSize As Integer = 4, Optional ByVal showFontOptions As Boolean = True) As String


            Dim sb As StringBuilder = New StringBuilder("")
            Dim strHiddenText As String = strFldName & "Text"
            Dim strHiddenTextPassed As String = strFldName & "Val"
            Dim strPostedVal As String = strFldName & "FilteredHTML"
            Dim strVarText As String = strFldName & "Var"
            Dim strVarTextString As String = strFldName & "VarString"

            _HasDHTMLControls = True

            'add this functionality to the submit form
            sb_dhtml.Append("var " & strPostedVal & " = new String(document." & strFldName & ".FilterSourceCode (document." & strFldName & ".DOM.body.innerHTML));" & vbCrLf)
            sb_dhtml.Append("document." & _form_name & "." & strHiddenText & ".value = " & strPostedVal & ";" & vbCrLf)

            'add this to windows_onload function
            sb_onload.Append("	var " & strVarText & " = document." & _form_name & "." & strHiddenTextPassed & ".value; " & vbCrLf)
            sb_onload.Append(strVarTextString & " = new String(" & strVarText & ") " & vbCrLf)
            sb_onload.Append("document." & strFldName & ".documentHTML = " & strVarTextString & ";" & vbCrLf)


            If strInstructions <> "" Then sb.Append(vbCrLf & "<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

            'add Menu
            sb.Append("<tr><td></td><td align=center valign=top  nowrap>" & TextAreaMenu(strFldName, strFonts, intFontMaxSize, showFontOptions) & "</td></tr>")

            sb.Append("<tr><td align=right valign=top  nowrap ")

            ''If blIsRequired Then sb.Append("class=frmRequired>*") Else 

            sb.Append("class=frmRequired>")

            sb.Append(strTitle & ":</td>")

            sb.Append("<td align=left valign=top class=frmField>")

            sb.Append("<table border=0 align=center valign=top cellpadding=0 cellspacing=0 width=" & intWidth & ">" & vbCrLf)

            'add Control
            sb.Append("<tr><td align=center valign=top colspan=3><object classid=clsid:2D360201-FFF5-11d1-8D03-00A0C959BC0A id=" & strFldName & " height=" & intHeight & " width=" & intWidth & " VIEWASTEXT></object></td></tr>" & vbCrLf)

            sb.Append("</table>" & vbCrLf)

            'this is the hidden value to pass to the control
            sb.Append("<input type=hidden name=" & strHiddenTextPassed & " id=" & strHiddenTextPassed & " value=""" & strFldVal & """>")

            'this is the hidden value to post
            sb.Append("<input type=hidden name=" & strHiddenText & " id=" & strHiddenText & ">")

            sb.Append("</td><tr>" & vbCrLf)

            'separator
            'sb.Append("<tr><td align=center valign=top colspan=2><hr noshade></td></tr>")


            Return sb.ToString()

        End Function


        Private Function DHTMLTextAreaJavaScript() As String

            Dim sb As StringBuilder = New StringBuilder("")

            sb.Append("//DHTML Script " & vbCrLf)

            sb.Append("// Command IDs" & vbCrLf)
            sb.Append("//" & vbCrLf)
            sb.Append("DECMD_BOLD =                      5000" & vbCrLf)
            sb.Append("DECMD_COPY =                      5002" & vbCrLf)
            sb.Append("DECMD_CUT =                       5003" & vbCrLf)
            sb.Append("DECMD_DELETE =                    5004" & vbCrLf)
            sb.Append("DECMD_DELETECELLS =               5005" & vbCrLf)
            sb.Append("DECMD_DELETECOLS =                5006" & vbCrLf)
            sb.Append("DECMD_DELETEROWS =                5007" & vbCrLf)
            sb.Append("DECMD_FINDTEXT =                  5008" & vbCrLf)
            sb.Append("DECMD_FONT =                      5009" & vbCrLf)
            sb.Append("DECMD_GETBACKCOLOR =              5010" & vbCrLf)
            sb.Append("DECMD_GETBLOCKFMT =               5011" & vbCrLf)
            sb.Append("DECMD_GETBLOCKFMTNAMES =          5012" & vbCrLf)
            sb.Append("DECMD_GETFONTNAME =               5013" & vbCrLf)
            sb.Append("DECMD_GETFONTSIZE =               5014" & vbCrLf)
            sb.Append("DECMD_GETFORECOLOR =              5015" & vbCrLf)
            sb.Append("DECMD_HYPERLINK =                 5016" & vbCrLf)
            sb.Append("DECMD_IMAGE =                     5017" & vbCrLf)
            sb.Append("DECMD_INDENT =                    5018" & vbCrLf)
            sb.Append("DECMD_INSERTCELL =                5019" & vbCrLf)
            sb.Append("DECMD_INSERTCOL =                 5020" & vbCrLf)
            sb.Append("DECMD_INSERTROW =                 5021" & vbCrLf)
            sb.Append("DECMD_INSERTTABLE =               5022" & vbCrLf)
            sb.Append("DECMD_ITALIC =                    5023" & vbCrLf)
            sb.Append("DECMD_JUSTIFYCENTER =             5024" & vbCrLf)
            sb.Append("DECMD_JUSTIFYLEFT =               5025" & vbCrLf)
            sb.Append("DECMD_JUSTIFYRIGHT =              5026" & vbCrLf)
            sb.Append("DECMD_LOCK_ELEMENT =              5027" & vbCrLf)
            sb.Append("DECMD_MAKE_ABSOLUTE =             5028" & vbCrLf)
            sb.Append("DECMD_MERGECELLS =                5029" & vbCrLf)
            sb.Append("DECMD_ORDERLIST =                 5030" & vbCrLf)
            sb.Append("DECMD_OUTDENT =                   5031" & vbCrLf)
            sb.Append("DECMD_PASTE =                     5032" & vbCrLf)
            sb.Append("DECMD_REDO =                      5033" & vbCrLf)
            sb.Append("DECMD_REMOVEFORMAT =              5034" & vbCrLf)
            sb.Append("DECMD_SELECTALL =                 5035" & vbCrLf)
            sb.Append("DECMD_SEND_BACKWARD =             5036" & vbCrLf)
            sb.Append("DECMD_BRING_FORWARD =             5037" & vbCrLf)
            sb.Append("DECMD_SEND_BELOW_TEXT =           5038" & vbCrLf)
            sb.Append("DECMD_BRING_ABOVE_TEXT =          5039" & vbCrLf)
            sb.Append("DECMD_SEND_TO_BACK =              5040" & vbCrLf)
            sb.Append("DECMD_BRING_TO_FRONT =            5041" & vbCrLf)
            sb.Append("DECMD_SETBACKCOLOR =              5042" & vbCrLf)
            sb.Append("DECMD_SETBLOCKFMT =               5043" & vbCrLf)
            sb.Append("DECMD_SETFONTNAME =               5044" & vbCrLf)
            sb.Append("DECMD_SETFONTSIZE =               5045" & vbCrLf)
            sb.Append("DECMD_SETFORECOLOR =              5046" & vbCrLf)
            sb.Append("DECMD_SPLITCELL =                 5047" & vbCrLf)
            sb.Append("DECMD_UNDERLINE =                 5048" & vbCrLf)
            sb.Append("DECMD_UNDO =                      5049" & vbCrLf)
            sb.Append("DECMD_UNLINK =                    5050" & vbCrLf)
            sb.Append("DECMD_UNORDERLIST =               5051" & vbCrLf)
            sb.Append("DECMD_PROPERTIES =                5052" & vbCrLf)
            sb.Append("//" & vbCrLf)
            sb.Append("// Enums" & vbCrLf)
            sb.Append("//" & vbCrLf)

            sb.Append("// OLECMDEXECOPT  " & vbCrLf)
            sb.Append("OLECMDEXECOPT_DODEFAULT =         0 " & vbCrLf)
            sb.Append("OLECMDEXECOPT_PROMPTUSER =        1" & vbCrLf)
            sb.Append("OLECMDEXECOPT_DONTPROMPTUSER =    2" & vbCrLf)
            sb.Append("" & vbCrLf)
            sb.Append("// DHTMLEDITCMDF" & vbCrLf)
            sb.Append("DECMDF_NOTSUPPORTED =             0 " & vbCrLf)
            sb.Append("DECMDF_DISABLED =                 1 " & vbCrLf)
            sb.Append("DECMDF_ENABLED =                  3" & vbCrLf)
            sb.Append("DECMDF_LATCHED =                  7" & vbCrLf)
            sb.Append("DECMDF_NINCHED =                  11" & vbCrLf)
            sb.Append("" & vbCrLf)
            sb.Append("// DHTMLEDITAPPEARANCE" & vbCrLf)
            sb.Append("DEAPPEARANCE_FLAT =               0" & vbCrLf)
            sb.Append("DEAPPEARANCE_3D =                 1 " & vbCrLf)
            sb.Append("" & vbCrLf)
            sb.Append("// OLE_TRISTATE" & vbCrLf)
            sb.Append("OLE_TRISTATE_UNCHECKED =          0" & vbCrLf)
            sb.Append("OLE_TRISTATE_CHECKED =            1" & vbCrLf)
            sb.Append("OLE_TRISTATE_GRAY =               2" & vbCrLf)

            sb.Append("" & vbCrLf)
            sb.Append("" & vbCrLf)

            'images
            sb.Append("//Images and rollovers " & vbCrLf)
            sb.Append("if (document.images) {" & vbCrLf)
            sb.Append("	bold 		= new Image;" & vbCrLf)
            sb.Append("	bold_o 		= new Image;" & vbCrLf)
            sb.Append("	under 		= new Image;" & vbCrLf)
            sb.Append("	under_o		= new Image;" & vbCrLf)
            sb.Append("	right 		= new Image;" & vbCrLf)
            sb.Append("	right_o		= new Image;" & vbCrLf)
            sb.Append("	center 		= new Image;" & vbCrLf)
            sb.Append("	center_o	= new Image;" & vbCrLf)
            sb.Append("	left 		= new Image;" & vbCrLf)
            sb.Append("	left_o 		= new Image;" & vbCrLf)
            sb.Append("	italic 		= new Image;" & vbCrLf)
            sb.Append("	italic_o	= new Image;" & vbCrLf)
            sb.Append("	link 		= new Image;" & vbCrLf)
            sb.Append("	link_o 		= new Image;" & vbCrLf)
            sb.Append("	numlist		= new Image;" & vbCrLf)
            sb.Append("	numlist_o	= new Image;" & vbCrLf)
            sb.Append("	bullist		= new Image;" & vbCrLf)
            sb.Append("	bullist_o	= new Image;" & vbCrLf)
            sb.Append("	inindent	= new Image;" & vbCrLf)
            sb.Append("	inindent_o	= new Image;" & vbCrLf)
            sb.Append("	deindent	= new Image;" & vbCrLf)
            sb.Append("	deindent_o	= new Image;" & vbCrLf)
            sb.Append("	" & vbCrLf)
            sb.Append("	bold.src 		= """ & _project_path & "images/bold.gif"";" & vbCrLf)
            sb.Append("	bold_o.src 		= """ & _project_path & "images/bold_o.gif"";" & vbCrLf)
            sb.Append("	under.src 		= """ & _project_path & "images/under.gif"";" & vbCrLf)
            sb.Append("	under_o.src 	= """ & _project_path & "images/under_o.gif"";" & vbCrLf)
            sb.Append("	right.src 		= """ & _project_path & "images/right.gif"";" & vbCrLf)
            sb.Append("	right_o.src 	= """ & _project_path & "images/right_o.gif"";" & vbCrLf)
            sb.Append("	center.src 		= """ & _project_path & "images/center.gif"";" & vbCrLf)
            sb.Append("	center_o.src	= """ & _project_path & "images/center_o.gif"";" & vbCrLf)
            sb.Append("	left.src 		= """ & _project_path & "images/left.gif"";" & vbCrLf)
            sb.Append("	left_o.src 		= """ & _project_path & "images/left_o.gif"";" & vbCrLf)
            sb.Append("	italic.src 		= """ & _project_path & "images/italic.gif"";" & vbCrLf)
            sb.Append("	italic_o.src	= """ & _project_path & "images/italic_o.gif"";" & vbCrLf)
            sb.Append("	link.src 		= """ & _project_path & "images/link.gif"";" & vbCrLf)
            sb.Append("	link_o.src 		= """ & _project_path & "images/link_o.gif"";" & vbCrLf)
            sb.Append("	numlist.src 	= """ & _project_path & "images/numlist.gif"";" & vbCrLf)
            sb.Append("	numlist_o.src 	= """ & _project_path & "images/numlist_o.gif"";" & vbCrLf)
            sb.Append("	bullist.src 	= """ & _project_path & "images/bullist.gif"";" & vbCrLf)
            sb.Append("	bullist_o.src 	= """ & _project_path & "images/bullist_o.gif"";" & vbCrLf)
            sb.Append("	inindent.src 	= """ & _project_path & "images/inindent.gif"";" & vbCrLf)
            sb.Append("	inindent_o.src 	= """ & _project_path & "images/inindent_o.gif"";" & vbCrLf)
            sb.Append("	deindent.src 	= """ & _project_path & "images/deindent.gif"";" & vbCrLf)
            sb.Append("	deindent_o.src 	= """ & _project_path & "images/deindent_o.gif"";" & vbCrLf)
            sb.Append("	" & vbCrLf)
            sb.Append("}" & vbCrLf)

            sb.Append("//Script Functions and Event Handlers " & vbCrLf)

            sb.Append("function DECMD_UNORDERLIST_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_UNORDERLIST,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_UNDERLINE_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_UNDERLINE,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_OUTDENT_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_OUTDENT,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_ORDERLIST_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_ORDERLIST,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_JUSTIFYRIGHT_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_JUSTIFYRIGHT,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_JUSTIFYLEFT_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_JUSTIFYLEFT,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_JUSTIFYCENTER_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_JUSTIFYCENTER,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_ITALIC_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_ITALIC,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_INDENT_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_INDENT,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_HYPERLINK_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_HYPERLINK,OLECMDEXECOPT_PROMPTUSER);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)

            sb.Append("function DECMD_BOLD_onclick(callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_BOLD,OLECMDEXECOPT_DODEFAULT);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)
            sb.Append("" & vbCrLf)


            sb.Append("function FontSize_onchange1(MyFontSize, callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_SETFONTSIZE, OLECMDEXECOPT_DODEFAULT, parseInt(MyFontSize));" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)


            sb.Append("function FontName_onchange1(MyFont, callingElement) {" & vbCrLf)
            sb.Append("callingElement.ExecCommand(DECMD_SETFONTNAME, OLECMDEXECOPT_DODEFAULT, MyFont);" & vbCrLf)
            sb.Append("callingElement.focus();" & vbCrLf)
            sb.Append("}" & vbCrLf)

            'window_onload function

            sb.Append(OnLoadScript())

            Return sb.ToString()

        End Function


        Private Function TextAreaMenu(ByVal strName As String, Optional ByVal strFonts As String = "Arial,Roman,Verdana", Optional ByVal intMaxSize As Integer = 4, Optional ByVal showFontOptions As Boolean = True) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim intIndex As Integer = 0
            Dim arrFonts() As String

            sb.Append("<table border=0 cellpadding=0 cellspacing=0>")

            sb.Append("<tr class=frmLabel>" & vbCrLf)

            sb.Append("<td align=center valign=top nowrap>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_BOLD_onclick(document." & strName & ")"" onMouseOver=""imageRollover('boldpic" & strName & "', 'bold_o')"" onMouseOut=""imageRollover('boldpic" & strName & "','bold')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/bold.gif"" border=""0"" name=""boldpic" & strName & """ title=""Bold"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)


            sb.Append("<a href=""javascript:DECMD_ITALIC_onclick(document." & strName & ")"" onMouseOver=""imageRollover('italicpic" & strName & "', 'italic_o')"" onMouseOut=""imageRollover('italicpic" & strName & "','italic')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/italic.gif"" border=""0"" name=""italicpic" & strName & """ title=""Italic"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_UNDERLINE_onclick(document." & strName & ")"" onMouseOver=""imageRollover('underpic" & strName & "', 'under_o')"" onMouseOut=""imageRollover('underpic" & strName & "','under')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/under.gif"" border=""0"" name=""underpic" & strName & """ title=""Underline"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_OUTDENT_onclick(document." & strName & ")"" onMouseOver=""imageRollover('deindentpic" & strName & "', 'deindent_o')"" onMouseOut=""imageRollover('deindentpic" & strName & "','deindent')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/deindent.gif"" border=""0"" name=""deindentpic" & strName & """ title=""Outdent"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_INDENT_onclick(document." & strName & ")"" onMouseOver=""imageRollover('inindentpic" & strName & "', 'inindent_o')"" onMouseOut=""imageRollover('inindentpic" & strName & "','inindent')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/inindent.gif"" border=""0"" name=""inindentpic" & strName & """ title=""Indent"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_JUSTIFYLEFT_onclick(document." & strName & ")"" onMouseOver=""imageRollover('leftpic" & strName & "', 'left_o')"" onMouseOut=""imageRollover('leftpic" & strName & "','left')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/left.gif"" border=""0"" name=""leftpic" & strName & """ title=""Left Justify"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_JUSTIFYCENTER_onclick(document." & strName & ")"" onMouseOver=""imageRollover('centerpic" & strName & "', 'center_o')"" onMouseOut=""imageRollover('centerpic" & strName & "','center')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/center.gif"" border=""0"" name=""centerpic" & strName & """ title=""Center Justify"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_JUSTIFYRIGHT_onclick(document." & strName & ")"" onMouseOver=""imageRollover('rightpic" & strName & "', 'right_o')"" onMouseOut=""imageRollover('rightpic" & strName & "','right')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/right.gif"" border=""0"" name=""rightpic" & strName & """ title=""Right Justify"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_ORDERLIST_onclick(document." & strName & ")"" onMouseOver=""imageRollover('numlistpic" & strName & "', 'numlist_o')"" onMouseOut=""imageRollover('numlistpic" & strName & "','numlist')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/numlist.gif"" border=""0"" name=""numlistpic" & strName & """ title=""Ordered List"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

            sb.Append("<a href=""javascript:DECMD_UNORDERLIST_onclick(document." & strName & ")"" onMouseOver=""imageRollover('bullistpic" & strName & "', 'bullist_o')"" onMouseOut=""imageRollover('bullistpic" & strName & "','bullist')"">" & vbCrLf)
            sb.Append("<img src=""" & _project_path & "images/bullist.gif"" border=""0"" name=""bullistpic" & strName & """ title=""Unordered List"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)


            If showFontOptions Then

                sb.Append("<a href=""javascript:DECMD_HYPERLINK_onclick(document." & strName & ")"" onMouseOver=""imageRollover('linkpic" & strName & "', 'link_o')"" onMouseOut=""imageRollover('linkpic" & strName & "','link')"">" & vbCrLf)
                sb.Append("<img src=""" & _project_path & "images/link.gif"" border=""0"" name=""linkpic" & strName & """ title=""Hyperlink"" WIDTH=""23"" HEIGHT=""22""></a>" & vbCrLf)

                arrFonts = Split(strFonts, ",")

                'font size
                sb.Append("&nbsp;<b>Font Size:</b>" & vbCrLf)
                sb.Append("<select class=frmSelect ID=""FontSize" & strName & """ name=""FontSize" & strName & """ TITLE=""Font Size"" LANGUAGE=""javascript"" onchange=""return FontSize_onchange1(document." & _form_name & ".FontSize" & strName & ".value, document." & strName & ")"">" & vbCrLf)


                For intIndex = 1 To intMaxSize

                    sb.Append("<option value=""" & intIndex & """>" & intIndex & "</option>" & vbCrLf)

                Next

                sb.Append("</select>" & vbCrLf)

                'font
                sb.Append("&nbsp;<b>Font:</b>" & vbCrLf)
                sb.Append("<select class=frmSelect ID=""FontName" & strName & """ name=""FontName" & strName & """ TITLE=""Font Name"" LANGUAGE=""javascript"" onchange=""return FontName_onchange1(document." & _form_name & ".FontName" & strName & ".value, document." & strName & ")"">" & vbCrLf)

                For intIndex = LBound(arrFonts) To UBound(arrFonts)

                    sb.Append("<option value=""" & arrFonts(intIndex) & """>" & arrFonts(intIndex) & "</option>" & vbCrLf)

                Next

                sb.Append("</select>" & vbCrLf)

            End If

            'end cell, row, table
            sb.Append("</td>" & vbCrLf)

            sb.Append("</tr>" & vbCrLf)

            sb.Append("</table>")

            Return sb.ToString()

        End Function


        Private Function OnLoadScript() As String

            Dim sb As StringBuilder = New StringBuilder("")

            sb.Append("function window_onload() {" & vbCrLf)

            sb.Append("//edits the content of the field body so it will be in one line" & vbCrLf)

            sb.Append(sb_onload.ToString())

            sb.Append("}" & vbCrLf)


            Return sb.ToString()


        End Function


        Private Function FieldLayout(ByVal strLabel As String, ByVal strFrmFld As String, ByVal strInstructions As String, Optional ByVal isRequired As Boolean = False) As String

            Dim sb As New System.Text.StringBuilder("")

            Try

                If strInstructions <> "" Then sb.Append(vbCrLf & "<tr class=frmInst><td></td><td>" & strInstructions & "</td></tr>" & vbCrLf)

                sb.Append("<tr>")

                'form label
                sb.Append(vbCrLf & "<td align=right valign=top  nowrap ")
                If isRequired Then sb.Append("class=frmRequired>*") Else sb.Append("class=frmLabel>")
                sb.Append(strLabel & ":</td>")

                'form field
                sb.Append("<td align=left valign=top class=frmField>" & strFrmFld & "</td>")

                sb.Append("</tr>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Private Sub LoadXml()

            Dim doc As New System.Xml.XmlDocument
            Dim labels As System.Xml.XmlNodeList
            Dim strXPath As String = "/labels/webforms/webform"


            Try

                doc.Load(HttpContext.Current.Request.PhysicalApplicationPath & "\Editor\labels.xml")

                labels = doc.SelectNodes(strXPath)

                If labels.Count > 0 Then

                    For Each webform As System.Xml.XmlNode In labels

                        If webform.Attributes.GetNamedItem("name").Value = _form_name Then

                            _xml = webform

                        End If

                    Next

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Private Function FieldValuesFromXml(ByVal strFrmFldName As String) As String()

            Dim arrValues(1) As String

            Try

                HttpContext.Current.Trace.Write("before: If _xml.HasChildNodes Then")

                If IsNothing(_xml) Then

                    HttpContext.Current.Trace.Write("entered isnothing(_xml)")

                    Throw System.Web.HttpUnhandledException.CreateFromLastError("There are no labels for form: " & _form_name & " in file labels.xml.  Please make sure that the file contains the values")

                Else

                    If _xml.HasChildNodes Then

                        HttpContext.Current.Trace.Write("after: If _xml.HasChildNodes Then")

                        'Set values from xml
                        For Each fields As System.Xml.XmlNode In _xml

                            HttpContext.Current.Trace.Write("fields in _xml node: " & fields.Name & ": " & fields.Attributes("name").Value & " - " & fields.Attributes("label").Value & " - " & fields.Attributes("instructions").Value)

                            If fields.Attributes("name").Value = strFrmFldName Then

                                arrValues(0) = fields.Attributes("label").Value
                                arrValues(1) = fields.Attributes("instructions").Value

                            End If

                        Next

                        Return arrValues

                    Else

                        HttpContext.Current.Trace.Write("_xml.HasChildNodes returned false")

                        Throw System.Web.HttpUnhandledException.CreateFromLastError("There are no labels for " & _form_name & " in file labels.xml")

                    End If

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Function


        Public Function ShowXmlValues() As String

            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")

            For Each field As System.Xml.XmlNode In _xml

                sb.Append(field.Name & ": " & field.InnerText & "<br>")

            Next

            Return sb.ToString

        End Function


    End Class



End Namespace



