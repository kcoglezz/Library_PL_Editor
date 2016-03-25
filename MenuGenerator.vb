Imports System.Data
Imports System.Web

Namespace ParaLideres

    Public Class MenuGenerator

        Public Enum Direction As Integer

            Horizontal = 1
            Vertical = 2

        End Enum

        Private _strStyle As String = ""

        Private _strBackColor As String = "#B7D983"
        Private _strForeColor As String = "#D0F29C"
        Private _strBorderColor As String = "#B7D983"
        Private _strHoverColor As String = "#CCD1D4"

        Private _intMaxLevel As Integer = 2
        Private _intWidth As Integer = 200
        Private _intCellPadding As Integer = 4
        Private _intCellSpacing As Integer = 0
        Private _intDirection As Direction = Direction.Vertical

        Private _objTableMenu As DataTable

        Public Property BackColor() As String
            Get
                Return _strBackColor
            End Get
            Set(ByVal Value As String)
                _strBackColor = Value
            End Set
        End Property

        Public Property ForeColor() As String
            Get
                Return _strForeColor
            End Get
            Set(ByVal Value As String)
                _strForeColor = Value
            End Set
        End Property

        Public Property BorderColor() As String
            Get
                Return _strBorderColor
            End Get
            Set(ByVal Value As String)
                _strBorderColor = Value
            End Set
        End Property

        Public Property HoverColor() As String
            Get
                Return _strHoverColor
            End Get
            Set(ByVal Value As String)
                _strHoverColor = Value
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return _intWidth

            End Get
            Set(ByVal Value As Integer)

                _intWidth = Value

            End Set
        End Property

        Public Property MaxLevel() As Integer
            Get
                Return _intMaxLevel
            End Get
            Set(ByVal Value As Integer)
                _intMaxLevel = Value
            End Set
        End Property

        Public Property MenuDirection() As Direction
            Get
                Return _intDirection
            End Get
            Set(ByVal Value As Direction)

                _intDirection = Value

            End Set
        End Property
        Public Property CellPadding() As Integer
            Get
                Return _intCellPadding
            End Get
            Set(ByVal Value As Integer)

                _intCellPadding = Value

            End Set
        End Property
        Public Property CellSpacing() As Integer
            Get
                Return _intCellSpacing
            End Get
            Set(ByVal Value As Integer)

                _intCellSpacing = Value

            End Set
        End Property

        Public Function GenerateMenu(Optional ByVal strSection As String = "") As String

            Dim sb As New System.Text.StringBuilder("")

            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim arrMenuItems(6, 1) As String

            Dim strDesc As String = ""
            Dim strLink As String = ""
            Dim strSubMenu As String = ""

            Dim intSectionId As Integer = 0
            Dim intChildren As Integer = 0
            Dim intIndex As Integer = 0

            Dim isSelected As Boolean = False

            Try

                _strStyle = "style=""border-color:" & _strBorderColor & ";border-width:1px;border-style:solid;width:" & _intWidth & "px;border-collapse:separate;"""

                sb.Append("  <!-- MAIN MENU START -->" & Chr(13))

                sb.Append("<table cellpadding=" & _intCellPadding & " cellspacing=" & _intCellSpacing & " " & _strStyle & ">" & Chr(13))

                sb.Append("<tr><td align=center style=""background:#A7D168;color:white;text-decoration:none;font-family:Verdana;font-size:8pt;font-weight:bold;"">MEN&#250;</td></tr>")

                'GENERATE FROM DATABASE

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsForMenu 0, 1")

                If reader.HasRows() Then

                    Do While (reader.Read())

                        intSectionId = reader(0)
                        strDesc = reader(1)
                        intChildren = reader(2)

                        strLink = "section.aspx?section_id=" & intSectionId & "&index=0"

                        strSubMenu = SubMenu(intSectionId, 1)

                        'sb.Append(MenuRow(strLink, strDesc, intChildren, strSubMenu, isSelected))

                        MenuRow(strLink, strDesc, intChildren, strSubMenu, isSelected, True)

                    Loop

                End If

                'GENERATE FROM ARRAY

                'Private _arrLinks As String() = Split("||mi_cuenta.aspx|survey.aspx|buscar.aspx|busqueda_avanzada.aspx", "|")

                arrMenuItems(0, 0) = "lo_ultimo.aspx"
                arrMenuItems(0, 1) = "Lo &#250;ltimo"
                arrMenuItems(1, 0) = "cursos.aspx"
                arrMenuItems(1, 1) = "Cursos"
                arrMenuItems(2, 0) = "flash/files/historietas.htm"
                arrMenuItems(2, 1) = "Historietas"
                arrMenuItems(3, 0) = "forum.aspx"
                arrMenuItems(3, 1) = "Foro"
                arrMenuItems(4, 0) = "destacado.aspx"
                arrMenuItems(4, 1) = "Destacado"
                arrMenuItems(5, 0) = "mi_cuenta.aspx"
                arrMenuItems(5, 1) = "Mi Cuenta"
                arrMenuItems(6, 0) = "comments.aspx"
                arrMenuItems(6, 1) = "Sugerencias"

                For intIndex = 0 To 6

                    'sb.Append(MenuRow(arrMenuItems(intIndex, 0), arrMenuItems(intIndex, 1), 0, "", False))

                    MenuRow(arrMenuItems(intIndex, 0), arrMenuItems(intIndex, 1), 0, "", False, True)

                Next


                sb.Append(GetMenuItemsSorted())

                sb.Append("  <tr>" & Chr(13))
                sb.Append("   <td height=""1"" bgcolor=#CCCCCC></td>" & Chr(13))
                sb.Append("  </tr>" & Chr(13))

                sb.Append("  </table>" & Chr(13))
                sb.Append("  <!-- MAIN MENU END -->" & Chr(13))


                Return sb.ToString()

            Catch ex As Exception

                HttpContext.Current.Trace.Warn("MenuGenerator.GenerateMenu: " & ex.ToString())

                Throw ex

            Finally

                sb = Nothing

                reader.Close()
                reader = Nothing

            End Try

        End Function

        Private Function SubMenu(ByVal intParentId As Integer, ByVal intLevel As Integer) As String

            Dim sb As New System.Text.StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intSectionId As Integer
            Dim intChildren As Integer = 0

            Dim strSubMenu As String = ""
            Dim strDesc As String = ""
            Dim strLink As String = ""

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSectionsForMenu " & intParentId & ",1")

                If reader.HasRows() Then

                    intLevel = intLevel + 1

                    sb.Append("  <!-- SUB MENU START -->" & Chr(13))

                    sb.Append("<div name=SubMenu" & intParentId & " class=""menuNormal"" width=""" & _intWidth & """>" & Chr(13))

                    'sb.Append("<table border=""0"" cellspacing=""0"" cellpadding=""4"" bgcolor=""" & _strBackColor & """ width=""" & _intWidth & """ " & _strStyle & ">")

                    sb.Append("<table cellpadding=" & _intCellPadding & " cellspacing=" & _intCellSpacing & " " & _strStyle & ">" & Chr(13))

                    Do While (reader.Read())

                        intSectionId = reader(0)
                        strDesc = reader(1)
                        intChildren = reader(2)

                        strLink = "section.aspx?section_id=" & intSectionId & "&index=0"

                        If intLevel <= _intMaxLevel Then

                            strSubMenu = SubMenu(intSectionId, intLevel)

                        Else

                            strSubMenu = ""

                            intChildren = 0

                        End If

                        sb.Append(MenuRow(strLink, strDesc, intChildren, strSubMenu, False, False))

                    Loop

                    sb.Append("</table>" & Chr(13))
                    sb.Append("</div>" & Chr(13))

                    sb.Append("  <!-- SUB MENU END -->" & Chr(13))

                End If

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

                reader.Close()
                reader = Nothing

            End Try

        End Function

        Private Function MenuRow(ByVal strLink As String, ByVal strDesc As String, ByVal intChildren As Integer, ByVal strSubmenu As String, ByVal isSelected As Boolean, ByVal adToTable As Boolean) As String

            Dim sbContent As New System.Text.StringBuilder("")

            Dim strBackColor As String
            Dim strTRStyle As String
            Dim strHREFStyle As String
            Dim strStyle As String = "text-decoration:none;font-family:Verdana;font-size:8pt;font-weight:normal;"

            Try

                If isSelected Then

                    strBackColor = _strHoverColor

                Else

                    strBackColor = _strBackColor

                End If

                strTRStyle = "background:" & strBackColor & ";color:" & _strForeColor & ";" & strStyle

                strHREFStyle = "color:" & _strForeColor & ";" & strStyle

                If intChildren > 0 Then

                    'sbContent.Append("<tr style=""color:" & _strForeColor & ";text-decoration:none;"">" & vbCrLf)

                    sbContent.Append("<tr style=""" & strTRStyle & """>" & vbCrLf)

                    sbContent.Append("<td nowrap ")

                    sbContent.Append(" onmouseover=""")

                    sbContent.Append("expand(this, '" & _strHoverColor & "');")

                    sbContent.Append("""")

                    sbContent.Append(" onmouseout=""")

                    sbContent.Append("collapse(this, '" & strBackColor & "');")

                    sbContent.Append("""")

                    sbContent.Append(" >")

                    sbContent.Append("<a style=""" & strHREFStyle & """ href=""" & strLink & """>" & UCase(strDesc) & "</a>")

                    sbContent.Append(strSubmenu)

                    sbContent.Append("</td>" & vbLf)

                    sbContent.Append("</tr>" & vbLf)

                Else 'If intChildren > 0 Then

                    sbContent.Append("<tr style=""" & strTRStyle & """>" & vbCrLf)

                    sbContent.Append("<td nowrap ")

                    sbContent.Append(" onmouseover=""")

                    sbContent.Append("changeColor(this, '" & _strHoverColor & "');")

                    sbContent.Append("""")

                    sbContent.Append(" onmouseout=""")

                    sbContent.Append("changeColor(this, '" & strBackColor & "');")

                    sbContent.Append("""")

                    sbContent.Append(">")

                    sbContent.Append("<a style=""" & strHREFStyle & """ href=""" & strLink & """>" & UCase(strDesc) & "</a>")

                    'If intChildren > 0 Then sbContent.Append(strSubmenu)

                    sbContent.Append("</td>" & vbLf)

                    sbContent.Append("</tr>" & vbLf)

                End If 'If intChildren > 0 Then

                If adToTable Then AddMenuItem(strDesc, sbContent.ToString())

                Return sbContent.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sbContent = Nothing

            End Try

        End Function

        Private Sub AddMenuItem(ByVal strDesc As String, ByVal strContent As String)

            Dim objRow As DataRow

            Try

                objRow = _objTableMenu.NewRow()

                objRow("Description") = strDesc
                objRow("Value") = strContent

                _objTableMenu.Rows.Add(objRow)

            Catch ex As Exception

                Throw ex

            Finally

                objRow = Nothing

            End Try

        End Sub

        Private Function GetMenuItemsSorted() As String

            Dim dv As DataView
            Dim sb As New System.Text.StringBuilder("")

            Dim intCount As Integer = 0
            Dim intTop As Integer = 0

            Try

                dv = _objTableMenu.DefaultView

                dv.Sort = "Description"

                intTop = _objTableMenu.Rows.Count - 1

                For intCount = 0 To intTop

                    sb.Append(dv(intCount)("Value"))

                Next

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                dv = Nothing
                sb = Nothing

            End Try

        End Function


        Public Sub New()

            _objTableMenu = New DataTable("Menu")

            _objTableMenu.Columns.Add("Description", GetType(String))
            _objTableMenu.Columns.Add("Value", GetType(String))

        End Sub


        Protected Overrides Sub Finalize()

            _objTableMenu = Nothing

            MyBase.Finalize()

        End Sub
    End Class


End Namespace
