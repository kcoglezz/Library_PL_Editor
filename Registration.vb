Imports System.Text
'Imports DataLayer
Imports PLEditor.DataLayer
Imports EditorClass.DataLayer
Imports Library_PL_Editor.DataLayer

Namespace ParaLideres

    Public Class Registration

        Public Function RegForm(ByVal intUserId As Integer) As String

            Dim objFrm As New ParaLideres.FormControls.GenericForm("frmReg")
            Dim msb As StringBuilder = New StringBuilder("")
            Dim objRegisteredUser As reg_users

            Dim arrLenguas As String() = Split("Espa�ol|Portugu�s|Ingl�s|Otro", "|")

            Dim arrYesNo As String() = Split("No|Si", "|")

            Dim dtBday As Date = #1/1/1900#

            Dim blShowAll As Boolean = False

            Dim intCountryId As Integer = 0

            Try

                objRegisteredUser = New reg_users(intUserId)


                If intUserId = 0 Then

                    Try
                        intCountryId = CInt(ParaLideres.GenericDataHandler.ExecScalar("SELECT id FROM countries WHERE country_code = '" & Functions.GetCountryCode() & "'"))
                    Catch ex As Exception
                    End Try

                    objRegisteredUser.setCountry(intCountryId)

                    msb.Append("<p>Bienvenido, y gracias por tu inter�s en la juventud!</p>")
                    msb.Append("<p>Para poder registrarte necesitas darnos tu primer nombre, apellido, y tu e-mail. Debes estar seguro de que tu e-mail esta correcto ya que te vamos a enviar un mensaje con tu nombre de usuario y clave a esa direcci�n.</p>")
                    msb.Append("<p>Como usuario registrado podras publicar y bajar materiales y adem�s podr�s tener acceso a otras secciones de ParaL�deres.</p>")

                Else

                    dtBday = objRegisteredUser.getBday

                    blShowAll = True

                    msb.Append("<p>Bienvenido " & objRegisteredUser.getFirstName & " " & objRegisteredUser.getLastName & "!</p>")

                    msb.Append("<p>Aqu� podr�s actualizar tu informaci�n personal</p>")

                    msb.Append("<p>Recuerda que como usuario registrado podr�s publicar y bajar materiales y adem�s podr�s tener acceso a otras secciones de ParaL�deres.</p>")

                End If

                msb.Append(objFrm.FormAction("post_reg_user.aspx", 450))
                msb.Append(objFrm.FormHidden("reg_user_id", CStr(intUserId)))

                'FirstName
                msb.Append(objFrm.FormTextBox("Nombre(s)", "FirstName", objRegisteredUser.getFirstName, 50, "Inserte su nombre(s) (Luis, Luis Miguel, Maria, Maria Elena, Ma. Elena, etc.)", True))

                'LastName
                msb.Append(objFrm.FormTextBox("Apellido(s)", "LastName", objRegisteredUser.getLastName, 50, "Inserte su apellido(s) (Le�n, Le�n Gonzalez, etc.)", True))

                If intUserId > 0 Then

                    'Password
                    msb.Append(objFrm.FormPasswordTextBox("Clave Secreta", "Password", objRegisteredUser.getPassword, 16, "Inserte su clave secreta. M�nimo 4 letras y m�ximo 16 letras.", True, 4, False))

                End If

                'Email
                msb.Append(objFrm.FormEmailVerifyBox("Email", "Email", objRegisteredUser.getEmail, "Inserte su direcci�n de E-mail"))

                'City
                msb.Append(objFrm.FormTextBox("Ciudad", "City", objRegisteredUser.getCity, 50, "La ciudad donde vives", True))

                'State
                msb.Append(objFrm.FormTextBox("Estado", "State", objRegisteredUser.getState, 100, "El estado/provincia donde vives", False))

                'Country
                msb.Append(objFrm.FormSelect("Pa�s", "Country", objRegisteredUser.getCountry, "sp_GetAllCountries", "Selecciona el pa�s donde vives", True))

                'ShowInfo
                msb.Append(objFrm.FormSelectArray("Quieres P�blicar tu Perf�l en el Directorio?", "ShowInfo", objRegisteredUser.getShowInfo, arrYesNo, "Selecciona si quieres que tu perfil aparezca en el directorio para que otros l�deres puedan contactarte.", False, True))

                'ReceiveEmails
                msb.Append(objFrm.FormSelectArray("Deseas Recibir Noticias de Para L�deres?", "ReceiveEmails", objRegisteredUser.getReceiveEmails, arrYesNo, "Selecciona si quieres recibir noticias y otros anuncios de Para L�deres.", False, True))

                If blShowAll Then

                    'Bday
                    'msb.Append(objFrm.FormDateCal("Fecha de Nacimiento", "Bday", dtBday, "Seleccionar D�a de Nacimiento", True))
                    msb.Append(objFrm.FormDate("Fecha de Nacimiento", "Bday", dtBday, "Seleccionar D�a de Nacimiento", True, #1/1/1940#, Date.Today))

                    'Sex
                    msb.Append(objFrm.FormSelect("Sexo", "Sex", objRegisteredUser.getSex, "sp_GetAllSexo", "Selecciona Masculino o Femenino", True))

                    'MStatus
                    msb.Append(objFrm.FormSelect("Estado Civil", "MStatus", objRegisteredUser.getMStatus, "sp_GetAllEstadoCivil", "Selecciona tu estado civil", True))

                    'WorkType
                    msb.Append(objFrm.FormSelect("Tipo de Trabajo", "WorkType", objRegisteredUser.getWorkType, "sp_GetAllWorkTypes", "Selecciona tu tipo de trabajo con j�venes", True))

                    'MainLanguage
                    msb.Append(objFrm.FormSelectArray("Idioma Principal", "MainLanguage", objRegisteredUser.getMainLanguage, arrLenguas, "Selecciona tu lenguage principal", True, False))

                    'Phone
                    msb.Append(objFrm.FormTextBox("Tel�fono", "Phone", objRegisteredUser.getPhone, 15, "Tu n�mero de tel�fono", False))

                    'Picture
                    'msb.Append(objFrm.FormTextBox("Picture", "Picture", objRegisteredUser.getPicture, 50, "Enter Picture", True))

                    'Otherinfo
                    msb.Append(objFrm.FormTextArea("Perf�l", "Otherinfo", objRegisteredUser.getOtherinfo, 20, 70, "Ingresa informaci�n acerca de t� que quieras compartir con otros miembros de Para L�deres.", False, 400))

                End If  'If blShowAll Then

                msb.Append(objFrm.FormTextBox("N�mero", "verify_num", "", 5, "Por favor escribe el n�mero que se muestra en la imagen.", True, 5, True, False, True))

                msb.Append(objFrm.FormEnd())

                Return msb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                objRegisteredUser = Nothing
                objFrm = Nothing
                msb = Nothing

            End Try


        End Function

        Public Function ForgotEmailForm(ByVal strEmail As String) As String

            Dim objFrm As New ParaLideres.FormControls.GenericForm
            Dim msb As StringBuilder = New StringBuilder("")

            Try

                msb.Append(objFrm.FormAction("emailpassword.aspx", 500))
                msb.Append(objFrm.FormHidden("send", "yes"))

                'Email
                msb.Append(objFrm.FormTextBox("Email", "Email", strEmail, 100, "Ingresa la direcci�n de E-mail con la que te registraste", True, 5, False, True))

                msb.Append(objFrm.FormEnd())

                Return msb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                objFrm = Nothing
                msb = Nothing

            End Try

        End Function

        Public Function SendPassword(ByVal strEmail As String) As String

            Dim sb As New System.Text.StringBuilder("")

            Dim strPassword As String = ""

            Try

                If InStr(strEmail, "@") > 0 And InStr(strEmail, ".") > 0 Then

                    Try
                        strPassword = ParaLideres.GenericDataHandler.ExecScalar("proc_GetPasswordByEmail '" & strEmail & "'")
                    Catch ex As Exception
                        strPassword = ""
                    End Try

                    If strPassword <> "" Then

                        Functions.SendMail(Functions.SupportAccount, strEmail, "Informaci�n para ingresar a Para L�deres", "Tu clave para ingresar a Para L�deres (www.paralideres.org) es: " & strPassword)

                        sb.Append("Hemos enviado tu clave a la siguiente direcci�n: " & strEmail)

                    Else

                        sb.Append("No encontramos informaci�n que corresponda a la direcci�n <b>" & strEmail & "</b>.  Por favor verifica que la direcci�n de e-mail que proporcionaste es la correcta.")

                        sb.Append(ForgotEmailForm(""))

                    End If

                Else 'If InStr(strEmail, "@") > 0 And InStr(strEmail, ".") > 0 Then

                    sb.Append("El e-mail: " & strEmail & " no es una direcci�n de correo electr�nico v�lida.")

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

    End Class

End Namespace

