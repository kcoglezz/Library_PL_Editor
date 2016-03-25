Imports System.Net.Mail
Imports Sistema.PL.Entidad

Namespace ParaLideres
    Public Class Funcion
        'Private Shared _CredentialMail As String = System.Web.Configuration.WebConfigurationManager.AppSettings("CredentialMail")
        'Private Shared _CrendetialPass As String = System.Web.Configuration.WebConfigurationManager.AppSettings("CrendetialPass")
        'Private Shared _HabilitaSSL As String = System.Web.Configuration.WebConfigurationManager.AppSettings("HabilitaSSL")
        'Private Shared _PuertoMail As Integer = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings("PuertoMail"))
        'Private Shared _MailServer As Integer = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings("MailServer"))

        'Private Shared _CredentialMailOut As String = System.Web.Configuration.WebConfigurationManager.AppSettings("CredentialMailOut")
        'Private Shared _CrendetialPassOut As String = System.Web.Configuration.WebConfigurationManager.AppSettings("CrendetialPassOut")
        'Private Shared _HabilitaSSLOut As String = System.Web.Configuration.WebConfigurationManager.AppSettings("HabilitaSSLOut")
        'Private Shared _PuertoMailOut As Integer = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings("PuertoMailOut"))
        'Private Shared _MailServerOut As Integer = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings("MailServerOut"))



        Public Shared Function EnviarCorreo(ByVal DatosMail As InfoEnvioMail, ByVal ServerMail As InfoEMailServer) As Boolean
            Dim MailServer As String = ""
            Dim PuertoMail As String = ""
            Dim HabilitaSSL As String = ""
            Dim CredentialMail As String = ""
            Dim CrendetialPass As String = ""


            MailServer = ServerMail.MailServer
            PuertoMail = ServerMail.PuertoMail
            HabilitaSSL = ServerMail.HabilitaSSL
            CredentialMail = ServerMail.CredentialMail
            CrendetialPass = ServerMail.CrendetialPass
            

            Dim Mail As New MailMessage()
            Mail.From = New System.Net.Mail.MailAddress(DatosMail.De.ToString(), DatosMail.DesplieguedelNombre.ToString())
            Mail.To.Add(DatosMail.Para.ToString())
            If DatosMail.Cc.ToString() <> "" Then
                Mail.CC.Add(DatosMail.Cc.ToString())
            End If
            If DatosMail.Cco.ToString() <> "" Then
                Mail.Bcc.Add(DatosMail.Cco.ToString())
            End If
            Mail.Subject = DatosMail.Asunto.ToString()
            Mail.Body = DatosMail.Contenido.ToString()
            Mail.IsBodyHtml = DatosMail.EsHtml

            Mail.Priority = MailPriority.Normal


            'Dim cliente As New SmtpClient(System.Web.Configuration.WebConfigurationManager.AppSettings("MailServer"))
            Dim cliente As New SmtpClient(MailServer)
            cliente.Credentials = New System.Net.NetworkCredential(CredentialMail, CrendetialPass)
            'cliente.EnableSsl = _HabilitaSSL

            If Convert.ToInt16(PuertoMail) > 0 Then
                cliente.Port = Convert.ToInt16(PuertoMail.ToString)
            End If

            Try
                cliente.Send(Mail)
            Catch ex As Exception
                Throw ex
                EnviarCorreo = False
            End Try


            EnviarCorreo = True

        End Function
    End Class
End Namespace

