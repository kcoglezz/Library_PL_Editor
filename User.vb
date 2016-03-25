Namespace ParaLideres

    Public Class User

        Private objUser As DataLayer.reg_users


        Sub New()

            objUser = New DataLayer.reg_users(0)

        End Sub

        Sub New(ByVal intUserId As Integer)

            objUser = New DataLayer.reg_users(intUserId)

        End Sub

        Public Function Logon(ByVal strEmail As String, ByVal strPassword As String) As Integer

            Dim intUserId As Integer = 0

            Try

                Try

                    intUserId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_CheckLogon '" & strEmail & "','" & strPassword & "'"))

                Catch
                End Try

                Return intUserId

            Catch ex As Exception

                Throw ex

            End Try

        End Function

        Public Function LogonToEditor(ByVal strEmail As String, ByVal strPassword As String) As Integer

            Dim intUserId As Integer = 0

            Try

                intUserId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_CheckLogonToEditor '" & strEmail & "','" & strPassword & "'"))

                Return intUserId

            Catch ex As Exception

                Throw ex

            End Try

        End Function



    End Class

End Namespace
