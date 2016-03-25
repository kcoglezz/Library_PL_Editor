Imports System
Imports System.Collections.Specialized
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web

Namespace ParaLideres

    Public Class SecureQueryString

        Inherits NameValueCollection

        Private Const cryptoKey As String = "letmego"

        Private ReadOnly IV As Byte() = New Byte(7) {240, 3, 45, 29, 0, 76, 173, 59}

        Public _blIsDebugMode As String = System.Web.Configuration.WebConfigurationManager.AppSettings("IsDebugMode")


        Sub New()

        End Sub


        Sub New(ByVal encStr As String)

            Try

                encStr = Replace(encStr, " ", "+")

                HttpContext.Current.Trace.Write("encStr: " & encStr)

                Deserialize(Decrypt(encStr))

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public ReadOnly Property EncryptedString() As String

            Get

                Return HttpUtility.UrlEncode(Encrypt(Serialize()))

            End Get

        End Property

        Public Overrides Function ToString() As String

            Return EncryptedString

        End Function



        Private Function Encrypt(ByVal serQS As String) As String

            Dim buffer As Byte() = Encoding.ASCII.GetBytes(serQS)

            Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider

            Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider

            Try

                des.Key = md5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(cryptoKey))

                des.IV = IV

                Return Convert.ToBase64String(des.CreateEncryptor.TransformFinalBlock(buffer, 0, buffer.Length))

            Catch ex As Exception

                Throw ex

            End Try

        End Function

        Private Function Decrypt(ByVal encQS As String) As String

            Try

                Dim buffer As Byte() = Convert.FromBase64String(encQS)

                Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider

                Dim MD5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider

                des.Key = MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(cryptoKey))

                des.IV = IV

                Return Encoding.ASCII.GetString(des.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length()))

            Catch ex As CryptographicException

                HttpContext.Current.Trace.Warn(ex.Source())
                HttpContext.Current.Trace.Warn(ex.Message())
                HttpContext.Current.Trace.Warn(ex.tostring())

                Throw ex

            Catch ex As FormatException

                HttpContext.Current.Trace.Warn(ex.Source())
                HttpContext.Current.Trace.Warn(ex.Message())
                HttpContext.Current.Trace.Warn(ex.tostring())

                Throw ex

            End Try

        End Function



        Private Sub Deserialize(ByVal decQS As String)

            Dim nameValuePairs As String() = decQS.Split("&")

            Dim i As Integer

            Try

                If _blIsDebugMode Then HttpContext.Current.Trace.Write("Decoded query string values")

                For i = 0 To nameValuePairs.Length - 1

                    Dim nameValue As String() = nameValuePairs(i).Split("=")

                    If nameValue.Length = 2 Then

                        Me.Add(HttpContext.Current.Server.HtmlDecode(nameValue(0)), HttpContext.Current.Server.HtmlDecode(Replace(nameValue(1), "|and|", "&")))

                        'Me.Add(nameValue(0), nameValue(1))

                        If _blIsDebugMode Then HttpContext.Current.Trace.Write(nameValue(0) & " = " & nameValue(1))

                    End If

                Next


            Catch ex As Exception

                Throw ex

            End Try

        End Sub


        Private Function Serialize() As String

            Dim sb As StringBuilder
            Dim key As String

            Try

                sb = New StringBuilder("")

                For Each key In Me.AllKeys

                    sb.Append(key)

                    sb.Append("=")

                    sb.Append(HttpContext.Current.Server.HtmlEncode(Replace(Me(key), "&", "|and|")))

                    'sb.Append(Me(key))

                    sb.Append("&")

                Next key

                Return sb.ToString

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

    End Class

End Namespace
