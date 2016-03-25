Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.ComponentModel
Imports System.Configuration
Imports System.Drawing
Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
Imports Sistema.PL.Entidad


Namespace ParaLideres

    Public Class EditorPageGrid

        Inherits Control
        Implements INamingContainer

        Private _sql_string As String = ""
        Private _title As String = ""
        Private _width As Integer = 500
        Private _default_message As String = "No se pudo encontrar informaci&#243;n que coincida con tu b&#250;squeda."
        Private _index As Integer
        Private _page_size As Integer
        Private _page_URL As String = ""
        Private _strNext As String = ""
        Private _strPrev As String = ""
        Private _user_level As Integer = 0
        Private dgGrid As DataGrid = New DataGrid
        Private intRecords As Integer = 0
        Private intSectionID As Integer = 0
        Private intPages As Integer = 0
        Private intCurrentPage As Integer = 0
        Private intDataGridLines As Integer = 0

        Private colorHeader As Color = ColorTranslator.FromHtml("#B7D983")
        Private colorAlternate As Color = ColorTranslator.FromHtml("#D0F29C")
        Private strColorAlternate As String = "#D0F29C"

        Private mContext As System.Web.HttpContext = System.Web.HttpContext.Current
        'Private strServer As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Server")
        'Private strDatabase As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Database")
        'Private strUsername As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Username")
        'Private strPassword As String = System.Web.Configuration.WebConfigurationManager.AppSettings("Password")
        'Private strConnection As String = "User ID=" & strUsername & ";Password=" & strPassword & ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer & ";Network Library =dbmssocn"
        Private Shared oConex As New InfoConexion()
        Private Shared strConnection As String = oConex.StringConnection
        
        Public Sub New()

        End Sub

        Public Sub New(ByVal queryString As String)

            If Len(queryString) > 0 Then _sql_string = queryString

        End Sub

        Public Sub New(ByVal queryString As String, ByVal defaultMessage As String)

            If Len(queryString) > 0 Then _sql_string = queryString

            If Len(defaultMessage) > 0 Then _default_message = defaultMessage

        End Sub

        Public Sub New(ByVal queryString As String, ByVal defaultMessage As String, ByVal strTitle As String)

            If Len(queryString) > 0 Then _sql_string = queryString

            If Len(defaultMessage) > 0 Then _default_message = defaultMessage

            If Len(strTitle) > 0 Then _title = strTitle

        End Sub

        Public Property Query() As String
            Get
                Return _sql_string
            End Get
            Set(ByVal Value As String)
                _sql_string = Value
            End Set
        End Property

        Public Property GridWidth() As Integer
            Get
                Return _width
            End Get
            Set(ByVal Value As Integer)
                _width = Value
            End Set
        End Property


        Public Property Index() As Integer

            Get

                Return _index

            End Get

            Set(ByVal Value As Integer)

                If Value < 1 Then Value = 1
                _index = Value

            End Set

        End Property

        Public Property PagingSize() As Integer

            Get

                Return _page_size

            End Get

            Set(ByVal Value As Integer)

                If Value < 1 Then Value = 1
                _page_size = Value

            End Set

        End Property

        Public Property PageUrl() As String

            Get

                Return _page_URL

            End Get

            Set(ByVal Value As String)

                _page_URL = Value

            End Set

        End Property

        Public Property UserLevel() As Integer
            Get

                Return _user_level

            End Get

            Set(ByVal Value As Integer)

                _user_level = Value

            End Set
        End Property


        Public Property DefaultMessage() As String
            Get

                Return _default_message

            End Get

            Set(ByVal Value As String)

                _default_message = Value

            End Set

        End Property

        Public Property Title() As String
            Get

                Return _title

            End Get

            Set(ByVal Value As String)

                _title = Value

            End Set

        End Property

        Public Property SectionID() As Integer

            Get

                Return intSectionID

            End Get

            Set(ByVal Value As Integer)

                intSectionID = Value

            End Set

        End Property

        Public Property DataGridLines() As Integer

            Get

                Return intDataGridLines

            End Get

            Set(ByVal Value As Integer)

                intDataGridLines = Value

            End Set

        End Property

        Public Function RenderToString() As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim sw As StringWriter = New StringWriter(sb)
            Dim tw As HtmlTextWriter = New HtmlTextWriter(sw)

            Dim mTable As Table = New Table
            Dim mRow As TableRow
            Dim mCell As TableCell

            Dim btnNext As Button = New Button
            Dim btnPrev As Button = New Button

            btnNext.Text = "Siguiente >>"
            btnNext.CommandName = "Next"
            btnNext.CommandArgument = "btnNextCmd"
            btnNext.CssClass = "frmButton"

            btnPrev.Text = "<< Previo"
            btnPrev.CommandName = "Previous"
            btnPrev.CommandArgument = "btnPrevCmd"
            btnPrev.CssClass = "frmButton"

            dgGrid.Width = Unit.Pixel(_width)
            dgGrid.CellPadding = "2"
            dgGrid.CellSpacing = "0"

            Select Case intDataGridLines

                Case 0

                    dgGrid.GridLines = GridLines.None

                Case 1

                    dgGrid.GridLines = GridLines.Horizontal

                Case 2

                    dgGrid.GridLines = GridLines.Both

                Case 3

                    dgGrid.GridLines = GridLines.Vertical

            End Select

            dgGrid.HorizontalAlign = HorizontalAlign.Center

            dgGrid.HeaderStyle.BackColor = colorHeader
            dgGrid.HeaderStyle.ForeColor = Color.Black
            dgGrid.HeaderStyle.CssClass = "TABLEHEADER"
            dgGrid.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            dgGrid.HeaderStyle.VerticalAlign = VerticalAlign.Middle
            dgGrid.HeaderStyle.Wrap = False

            dgGrid.AlternatingItemStyle.BackColor = colorAlternate
            dgGrid.AlternatingItemStyle.HorizontalAlign = HorizontalAlign.Justify
            dgGrid.AlternatingItemStyle.VerticalAlign = VerticalAlign.Top

            'define colors for items on grid
            dgGrid.ItemStyle.CssClass = "GEN"
            dgGrid.ItemStyle.BackColor = Color.White
            dgGrid.ItemStyle.ForeColor = Color.Black
            dgGrid.ItemStyle.HorizontalAlign = HorizontalAlign.Justify
            dgGrid.ItemStyle.VerticalAlign = VerticalAlign.Top

            dgGrid.AutoGenerateColumns = True
            dgGrid.BorderWidth = Unit.Pixel(1)
            dgGrid.BorderColor = Color.Black

            Me.BindDataGrid(_sql_string)

            If dgGrid.Items.Count > 0 Then

                mTable.BorderWidth = Unit.Pixel(0)
                mTable.CellPadding = 2
                mTable.CellSpacing = 0

                If Len(_title) > 0 Then

                    mRow = New TableRow
                    mCell = New TableCell
                    mCell.CssClass = "DEFAULTFONT3"
                    mCell.HorizontalAlign = HorizontalAlign.Center
                    mCell.VerticalAlign = VerticalAlign.Top
                    mCell.Text = _title
                    mRow.Cells.Add(mCell)
                    mTable.Rows.Add(mRow)

                End If

                'add datagrid to cell and cell to row and row to table
                mRow = New TableRow
                mCell = New TableCell
                mCell.CssClass = "GEN"
                mCell.HorizontalAlign = HorizontalAlign.Center
                mCell.VerticalAlign = VerticalAlign.Top
                mCell.Width = Unit.Pixel(_width)
                mCell.Controls.Add(dgGrid)
                mRow.Cells.Add(mCell)
                mTable.Rows.Add(mRow)

                'row displaying menu links
                mRow = New TableRow
                mCell = New TableCell
                mCell.CssClass = "GEN"
                mCell.HorizontalAlign = HorizontalAlign.Center
                mCell.VerticalAlign = VerticalAlign.Top
                mCell.Text = _strPrev & _strNext
                mRow.Cells.Add(mCell)
                mTable.Rows.Add(mRow)


                'row displaying number of records
                mRow = New TableRow
                mCell = New TableCell
                mCell.Height = System.Web.UI.WebControls.Unit.Pixel(30)
                mCell.CssClass = "GEN"
                mCell.HorizontalAlign = HorizontalAlign.Center
                mCell.VerticalAlign = VerticalAlign.Top
                mCell.Text = "P&#225;gina " & intCurrentPage & " de " & intPages
                mRow.Cells.Add(mCell)
                mTable.Rows.Add(mRow)

                Me.Controls.Add(mTable)

            Else

                sb.Append(_default_message)

            End If


            Try

                Me.Render(tw)

            Catch exc As Exception

                sb.Append(exc.ToString)

            Finally

                sw = Nothing
                tw = Nothing

            End Try

            Return sb.ToString

        End Function

        Private Function ExecReader(ByVal strSQL As String) As SqlDataReader

            Dim cmd As SqlCommand
            Dim conn As SqlConnection

            Try

                conn = New SqlConnection(strConnection)

                Try

                    mContext.Trace.Write("Trying to open connection: ")
                    conn.Open()
                    mContext.Trace.Write("Open connection succeded!")

                Catch exc As Exception

                    mContext.Trace.Write("Connection could not be openned")
                    mContext.Trace.Write(exc.Source)
                    mContext.Trace.Write(exc.Message)
                    mContext.Trace.Write(exc.ToString())

                End Try

                mContext.Trace.Write("cmd = New SqlCommand(strSQL, conn)")

                cmd = New SqlCommand(strSQL, conn)

                mContext.Trace.Write("after: cmd = New SqlCommand(strSQL, conn)")

                Return cmd.ExecuteReader(CommandBehavior.CloseConnection)

            Catch ex As Exception

                mContext.Trace.Warn("ExecReader on Generic DataGrid: " & ex.Source() & "<br>" & ex.Message() & "<br>" & ex.ToString() & "<br>" & ex.InnerException.ToString())

            Finally

                conn = Nothing
                cmd = Nothing

            End Try


        End Function

        Public Sub BindDataGrid(ByVal sqlString As String)

            Dim mConn As New SqlConnection(strConnection)
            Dim mDR As SqlDataReader
            Dim cmd As SqlCommand
            Dim intNumberOfRecords As Integer
            Dim intItemIndex As Integer

            Try

                cmd = New SqlCommand(sqlString, mConn)

                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@StartRow", SqlDbType.Int))
                cmd.Parameters("@StartRow").Value = Me.Index
                cmd.Parameters.Add(New SqlParameter("@StopRow", SqlDbType.Int))
                cmd.Parameters("@StopRow").Value = Me.Index + Me.PagingSize - 1
                cmd.Parameters.Add(New SqlParameter("@UserLevel", SqlDbType.Int))
                cmd.Parameters("@UserLevel").Value = Me.UserLevel
                cmd.Parameters.Add(New SqlParameter("@TotalRecs", SqlDbType.Int))
                cmd.Parameters("@TotalRecs").Direction = ParameterDirection.Output
                cmd.Parameters.Add(New SqlParameter("@SectionID", SqlDbType.Int))
                cmd.Parameters("@SectionID").Value = intSectionID

            Catch exc As Exception

                mContext.Trace.Write("Generic Data Grid Error: " & exc.ToString())

            End Try

            Try

                mConn.Open()

            Catch e As SqlException

                mContext.Trace.Write("Connection failed to open on Generic Data Handler: " & e.ToString)

            End Try


            Try

                mDR = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                dgGrid.DataSource = mDR

                dgGrid.DataBind()

                intNumberOfRecords = dgGrid.Items.Count - 1

                intRecords = CInt(cmd.Parameters("@TotalRecs").Value)
                intPages = Math.Ceiling(intRecords / Me.PagingSize)

                If _index > intRecords Then

                    _index = intRecords - _page_size

                    BindDataGrid(_sql_string)

                End If

            Catch e As SqlException

                mContext.Trace.Write("Generic Data Handler Error (SQL Server): " & e.Message)

            Catch ex As Exception

                mContext.Trace.Write("Generic Data Handler Error (SQL Server): " & ex.Message)

            Finally

                'sets the code to add highlight feature on datagrid
                For intItemIndex = 0 To intNumberOfRecords

                    Dim e = New System.Web.UI.WebControls.DataGridItemEventArgs(dgGrid.Items.Item(intItemIndex))
                    dgGrid_ItemDataBound(dgGrid, e)

                Next

                mDR.Close()
                cmd = Nothing
                mConn = Nothing
                mDR = Nothing

            End Try

            intCurrentPage = CInt((_index + _page_size - 1) / _page_size)

            Dim strNumericMenu As String = ""
            Dim intX As Integer = 0

            For intX = 0 To intPages - 1

                If intCurrentPage = intX + 1 Then

                    strNumericMenu = strNumericMenu & " " & intX + 1 & " "

                Else

                    strNumericMenu = strNumericMenu & " <a href=" & _page_URL & "?index=" & (intX * _page_size) + 1 & "&section_id=" & intSectionID & ">" & intX + 1 & "</a> "

                End If



            Next



            If intPages > 1 Then

                Select Case intCurrentPage

                    Case Is <= 1

                        _strPrev = "<< Previo " & strNumericMenu
                        _strNext = "<a href=" & _page_URL & "?index=" & _index + _page_size & "&section_id=" & intSectionID & ">Siguiente >></a>"

                    Case Is >= intPages

                        _strPrev = "<a href=" & _page_URL & "?index=" & _index - _page_size & "&section_id=" & intSectionID & "><< Previo</a> " & strNumericMenu
                        _strNext = "Siguiente >>"

                    Case Else

                        _strPrev = "<a href=" & _page_URL & "?index=" & _index - _page_size & "&section_id=" & intSectionID & "><< Previo</a> " & strNumericMenu
                        _strNext = "<a href=" & _page_URL & "?index=" & _index + _page_size & "&section_id=" & intSectionID & ">Siguiente >></a>"

                End Select

            End If


        End Sub

        Private Sub dgGrid_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)

            If e.Item.ItemType = ListItemType.Item Then

                'change e.Item.Cells(1).Attributes.Add to change color of specified cell starting at index 0

                e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='lightyellow';this.style.color='red'")
                e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='white';this.style.color='black'")

            End If

            If e.Item.ItemType = ListItemType.AlternatingItem Then

                e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='lightyellow';this.style.color='red'")
                e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='" & strColorAlternate & "';this.style.color='black'")

            End If

        End Sub

    End Class

End Namespace
