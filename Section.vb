Imports System
Imports System.Web
Imports System.Web.Mail
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Namespace ParaLideres

    Public Class Section


        Private _objSection As DataLayer.sections
        Private _intSectionId As Integer = 0
        Private _project_path As String = System.Web.Configuration.WebConfigurationManager.AppSettings("ProjectPath")
        Private _strPageUrl As String = _project_path & "section.aspx" 'HttpContext.Current.Request.ServerVariables("URL")

        Public Sub New()

            _objSection = New DataLayer.sections(_intSectionId)

        End Sub

        Public Sub New(ByVal intSectionId As Integer)

            _objSection = New DataLayer.sections(intSectionId)
            _intSectionId = intSectionId

        End Sub

        Public ReadOnly Property SectionName() As String

            Get

                Return _objSection.getSectionName

            End Get

        End Property

        Public ReadOnly Property SectionPath() As String
            Get

                Return GetPath()

            End Get
        End Property

        Public ReadOnly Property TotalArticles() As Integer
            Get
                Return _objSection.getArticleCount

            End Get
        End Property

        Public Function GetParentPath(ByVal intSectionID As Integer) As String

            Dim strCacheVar As String = "SectionParentPath" & intSectionID

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim sb As StringBuilder = New StringBuilder("")
                Dim reader As System.Data.SqlClient.SqlDataReader

                Dim intParentId As Integer = 0

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetParentInfo " & intSectionID)

                    If reader.HasRows Then

                        reader.Read()

                        intParentId = reader(0)

                        sb.Append("<a href=""" & GetLink(intParentId, 1) & """>" & reader(1) & "</a>|")

                        sb.Append(GetParentPath(intParentId))

                    End If

                    HttpContext.Current.Cache.Insert(strCacheVar, sb.ToString())

                Catch ex As Exception

                    Throw ex

                Finally

                    reader.Close()
                    reader = Nothing

                    sb = Nothing

                End Try

            End If

            Return HttpContext.Current.Cache.Get(strCacheVar)

        End Function

        Public Function GetPath(Optional ByVal intIndex As Integer = 0) As String

            Dim sb As StringBuilder = New StringBuilder("")

            Dim arrPath() As String

            Dim intI As Integer = 0

            Try

                arrPath = Split(GetParentPath(_intSectionId), "|")

                For intI = UBound(arrPath) To 0 Step -1

                    sb.Append(arrPath(intI) & " | ")

                Next

                sb.Append("<a href=""" & GetLink(_intSectionId, intIndex) & """>" & _objSection.getSectionName & "</a>")

                Return sb.ToString()

            Catch ex As Exception

                Throw ex

            Finally

                sb = Nothing

            End Try

        End Function

        Private Function GetLink(ByVal intSectionId As Integer, ByVal intIndex As Integer) As String

            Dim strLink As String = _strPageUrl & "?section_id=" & intSectionId & "&index=" & intIndex

            Return strLink

        End Function

        Public Function DisplaySection(ByVal intIndex As Integer) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader

            Dim intTotalRecords As Integer = 0
            Dim intPageId As Integer
            Dim intPageRating As Integer

            Dim datePosted As Date

            Dim strPageTitle As String
            Dim strPageBlurb As String
            Dim strPageAuthorName As String
            Dim strCacheVar As String = ""



            Try

                'Author Info
                sb.Append(ParaLideres.Functions.DisplayAuthorInfo(_objSection.getUserId))

                'Show Section Text
                sb.Append("<p>" & Functions.ReplaceCR(_objSection.getPagetext) & "</p>")

                ''Show Sub Sections
                sb.Append(SubSectionInfoByParentId(_intSectionId))

                strCacheVar = "TotalRecordsBySection" & _intSectionId

                If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    intTotalRecords = Functions.CountArticles("proc_CountPostedArticlesBySectionId " & _intSectionId)

                    HttpContext.Current.Cache.Insert(strCacheVar, intTotalRecords)

                Else

                    intTotalRecords = HttpContext.Current.Cache.Get(strCacheVar)

                End If

                If intTotalRecords > 0 Then

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetpagesToDisplayBySection " & _intSectionId & "," & intIndex)

                    If reader.HasRows Then

                        sb.Append("<ul class=UL>")

                        Do While (reader.Read())

                            intPageRating = 0

                            If Not reader.IsDBNull(1) Then intPageId = reader(1)
                            If Not reader.IsDBNull(2) Then datePosted = reader(2)
                            If Not reader.IsDBNull(3) Then strPageTitle = reader(3)
                            If Not reader.IsDBNull(4) Then strPageBlurb = Functions.ReplaceCR(reader(4))
                            If Not reader.IsDBNull(5) Then intPageRating = reader(5)
                            If Not reader.IsDBNull(6) Then strPageAuthorName = reader(6)

                            sb.Append(Functions.DiplayArticleInfo(intPageId, strPageTitle, intPageRating, strPageAuthorName, datePosted, strPageBlurb, True, intIndex))

                        Loop

                        sb.Append("</ul>")

                    End If 'If reader.HasRows Then

                    sb.Append("<p align=center>" & BottomMenu(intTotalRecords, intIndex) & "</p>")


                End If 'If intTotalRecords > 0 Then

                Return sb.ToString()

            Catch ex As Exception

                HttpContext.Current.Trace.Warn("pages by section: " & ex.ToString)

                Throw ex

            Finally

                If Not IsNothing(reader) Then

                    If Not reader.IsClosed Then reader.Close()
                    reader = Nothing

                End If

                sb = Nothing

            End Try

        End Function


        Private Function SubSectionInfoByParentId(ByVal intParentId As Integer) As String

            Dim strCacheVar As String = "SubSenctions" & intParentId

            Try

                If IsNothing(HttpContext.Current.Cache(strCacheVar)) Then


                    Dim sb As StringBuilder = New StringBuilder("")
                    Dim reader As System.Data.SqlClient.SqlDataReader

                    Dim intChildId As Integer
                    Dim strChildName As String
                    Dim strChildText As String
                    Dim intChildRecordCount As Integer

                    Try

                        'Show Sub Sections
                        reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSubSectionsBySectionID " & _intSectionId)

                        If reader.HasRows Then

                            Do While (reader.Read())

                                intChildId = 0
                                intChildRecordCount = 0

                                If Not reader.IsDBNull(0) Then intChildId = reader(0)
                                If Not reader.IsDBNull(1) Then strChildName = reader(1)
                                If Not reader.IsDBNull(2) Then strChildText = Functions.ReplaceCR(reader(2))
                                If Not reader.IsDBNull(3) Then intChildRecordCount = reader(3)

                                If intChildId Then sb.Append("<p><img src=" & _project_path & "images/seccion.gif><a href=" & GetLink(intChildId, 0) & ">" & strChildName & "</a> (" & intChildRecordCount & ")<br>" & strChildText)

                            Loop

                        End If


                        HttpContext.Current.Cache.Insert(strCacheVar, sb.ToString())

                    Catch ex As Exception

                        Throw ex

                    Finally

                        reader.Close()
                        reader = Nothing

                        sb = Nothing

                    End Try

                End If 'If IsNothing(HttpContext.Current.Cache(strCacheVar)) Then

                Return HttpContext.Current.Cache.Get(strCacheVar)

            Catch ex As Exception

                Throw ex

            End Try

        End Function


        Public Function DisplaySectionHtml(ByVal intIndex As Integer) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intTotalRecords As Integer = 0

            'Dim strAuthorName As String
            'Dim intAuthorId As Integer
            'Dim strAuthorPic As String

            Dim intChildId As Integer
            Dim strChildName As String
            Dim strChildText As String
            Dim intChildRecordCount As Integer

            Dim intPageId As Integer
            Dim datePosted As Date
            Dim strPageTitle As String
            Dim strPageBlurb As String
            Dim intPageRating As Integer
            Dim strPageAuthorName As String


            'Dim objFile As System.IO.File

            'Author Info
            sb.Append(ParaLideres.Functions.DisplayAuthorInfoHtml(_objSection.getUserId))

            'Show Section Text
            sb.Append("<p>" & Functions.ReplaceCR(_objSection.getPagetext) & "</p>")

            'Show Sub Sections
            reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetSubSectionsBySectionID " & _intSectionId)


            Try

                If reader.HasRows Then

                    Do While (reader.Read())

                        intChildId = 0
                        intChildRecordCount = 0

                        If Not reader.IsDBNull(0) Then intChildId = reader(0)
                        If Not reader.IsDBNull(1) Then strChildName = reader(1)
                        If Not reader.IsDBNull(2) Then strChildText = reader(2)
                        If Not reader.IsDBNull(3) Then intChildRecordCount = reader(3)

                        intChildRecordCount = intChildRecordCount + GetTotalNoOfArticles(intChildId)

                        If intChildId Then sb.Append("<p><img src=" & _project_path & "images/seccion.gif><a href=" & _project_path & "section.aspx?section_id=" & intChildId & ">" & strChildName & "</a> (" & intChildRecordCount & ")<br>" & strChildText)

                    Loop

                End If

            Catch ex As Exception

                HttpContext.Current.Trace.Warn("sub sections: " & ex.ToString)

                Throw ex

            Finally

                reader.Close()

            End Try


            'gettotal articles for subsections
            '_intTotalArticles = GetTotalNoOfArticles(_objSection.getSectionId)


            'Show Articles By Section

            Try
                intTotalRecords = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_CountPostedArticlesBySectionId " & _intSectionId))
            Catch ex As Exception
            End Try


            If intTotalRecords > 0 Then

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetpagesToDisplayBySection " & _intSectionId & ",1,200")

                Try

                    If reader.HasRows Then

                        sb.Append("<ul class=UL>")

                        Do While (reader.Read())

                            intPageRating = 0

                            If Not reader.IsDBNull(1) Then intPageId = reader(1)
                            If Not reader.IsDBNull(2) Then datePosted = reader(2) Else datePosted = #1/1/1900#
                            If Not reader.IsDBNull(3) Then strPageTitle = reader(3) Else strPageTitle = ""
                            If Not reader.IsDBNull(4) Then strPageBlurb = Functions.ReplaceCR(reader(4)) Else strPageBlurb = ""
                            If Not reader.IsDBNull(5) Then intPageRating = reader(5) Else intPageRating = 0
                            If Not reader.IsDBNull(6) Then strPageAuthorName = reader(6) Else strPageAuthorName = ""

                            sb.Append(Functions.DiplayArticleInfo(intPageId, strPageTitle, intPageRating, strPageAuthorName, datePosted, strPageBlurb))

                        Loop

                        sb.Append("</ul>")

                    End If

                Catch ex As Exception

                    HttpContext.Current.Trace.Warn("pages by section: " & ex.ToString)

                    Throw ex

                Finally

                    reader.Close()

                End Try


                'sb.Append("<p align=center>" & BottomMenu(intTotalRecords, intIndex) & "</p>")

            End If

            Return sb.ToString()

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

                    strNumericMenu = strNumericMenu & " <a href=" & _strPageUrl & "?index=" & (intX * intPageSize) + 1 & "&section_id=" & _intSectionId & ">" & intX + 1 & "</a> "

                End If

            Next

            If intPages > 1 Then

                Select Case intCurrentPage

                    Case Is <= 1

                        _strPrev = "<< Previo " & strNumericMenu
                        _strNext = "<a href=" & _strPageUrl & "?index=" & intIndex + intPageSize & "&section_id=" & _intSectionId & ">Siguiente >></a>"

                    Case Is >= intPages

                        _strPrev = "<a href=" & _strPageUrl & "?index=" & intIndex - intPageSize & "&section_id=" & _intSectionId & "><< Previo</a> " & strNumericMenu
                        _strNext = "Siguiente >>"

                    Case Else

                        _strPrev = "<a href=" & _strPageUrl & "?index=" & intIndex - intPageSize & "&section_id=" & _intSectionId & "><< Previo</a> " & strNumericMenu
                        _strNext = "<a href=" & _strPageUrl & "?index=" & intIndex + intPageSize & "&section_id=" & _intSectionId & ">Siguiente >></a>"

                End Select

            End If

            Return _strPrev & " " & _strNext

        End Function

        Public Function GetTotalNoOfArticles(ByVal intSectionId As Integer) As Integer

            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim intTotal As Integer = 0
            Dim intThisSection As Integer = 0
            Dim intThisTotal As Integer = 0

            Try

                reader = ParaLideres.GenericDataHandler.GetRecords("proc_CountPagesBySectionId " & intSectionId)

                If reader.HasRows Then

                    Do While (reader.Read())

                        intThisSection = reader(0)
                        intThisTotal = reader(1)

                        intTotal = intTotal + intThisTotal + GetTotalNoOfArticles(intThisSection)

                    Loop

                End If

                Return intTotal

            Catch ex As Exception

                Throw ex

            Finally

                reader.Close()
                reader = Nothing

            End Try

        End Function

    End Class

End Namespace
