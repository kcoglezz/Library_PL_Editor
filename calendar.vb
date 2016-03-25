Option Explicit On 
Imports System
Imports System.Text

Namespace ParaLideres.FormControls

    Public Class Calendar

        Private dtToday As Date = Today()
        Private dtFirstDay As Date
        Private dtLastDay As Date
        Private dtMonthDay As Date
        Private strScript As String
        Private strMonthName As String
        Private strPreviousDate As String
        Private strNextDate As String
        Private intPreviousMonth As Integer
        Private intNextMonth As Integer
        Private intNextYear As Integer
        Private intPrevMonth As Integer
        Private intPrevYear As Integer
        Private intThisYear As Integer
        Private intThisMonth As Integer
        Private intThisDay As Integer
        Private intFirstWeekDay As Integer
        Private intLastMonthDate As Integer
        Private intNextMonthDate As Integer
        Private strFormFieldName As String
        Private strProjectPath As String = System.Configuration.ConfigurationSettings.AppSettings("ProjectPath")
        Private strTopColor As String = "#CCCCCC"
        Private strMonthNameColor As String = "#000000"
        Private intYearStart As Integer = Today.Year() - 10
        Private intYearEnd As Integer = Today.Year() + 10

        Public Property Script() As String
            Get

                Return strScript

            End Get

            Set(ByVal Value As String)

                strScript = Value

            End Set
        End Property

        Public WriteOnly Property TopColor() As String

            Set(ByVal Value As String)

                strTopColor = Value

            End Set

        End Property

        Public WriteOnly Property MonthNameColor() As String

            Set(ByVal Value As String)

                strMonthNameColor = Value

            End Set

        End Property

        Public Property FieldName() As String
            Get
                Return strFormFieldName

            End Get
            Set(ByVal Value As String)

                strFormFieldName = Value

            End Set
        End Property

        Public WriteOnly Property YearStart() As Integer

            Set(ByVal Value As Integer)

                intYearStart = Value

            End Set

        End Property

        Public WriteOnly Property YearEnd() As Integer

            Set(ByVal Value As Integer)

                intYearEnd = Value

            End Set

        End Property


        Public Function DisplayCalendar(ByVal dtDateSelected As Date) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim iRowNum As Integer = 0
            Dim iLoopDay As Integer = 0

            If dtDateSelected = CDate("1/1/1900") Then dtDateSelected = dtToday

            ' Get the name of the file executing this routine
            'sScript = Request.ServerVariables("SCRIPT_NAME")

            ' Get the month and year values
            intThisDay = Day(dtDateSelected)
            intThisMonth = Month(dtDateSelected)
            intThisYear = Year(dtDateSelected)

            ' Get the name of the month
            strMonthName = Functions.SpanishMonthName(intThisMonth)

            ' Compute the first and last dates in the month. We will need these in our query
            dtFirstDay = DateSerial(intThisYear, intThisMonth, 1)
            dtLastDay = DateSerial(intThisYear, intThisMonth + 1, 0)

            ' Compute the first week day of the month
            intFirstWeekDay = Weekday(dtFirstDay, vbSunday)

            ' Get the previous month and year. This is used to create the back image
            intPrevMonth = Month(DateSerial(intThisYear, intThisMonth - 1, 1))
            intPrevYear = Year(DateSerial(intThisYear, intThisMonth - 1, 1))

            ' Get the next month and year. This information is used to drive the forward image
            intNextMonth = Month(DateSerial(intThisYear, intThisMonth + 1, 1))
            intNextYear = Year(DateSerial(intThisYear, intThisMonth + 1, 1))

            ' Get the last day of previous month. Using this, find the first sunday of the first
            ' week in the month
            intLastMonthDate = Day(DateSerial(intThisYear, intThisMonth, 0)) - intFirstWeekDay + 2
            intNextMonthDate = 1


            strPreviousDate = intPrevMonth & "/1/" & intPrevYear
            strNextDate = intNextMonth & "/1/" & intNextYear

            ' Initialize the date variable with the first date in the first week of the month
            dtMonthDay = DateSerial(intThisYear, intThisMonth, vbSunday - intFirstWeekDay + 1)

            'Stylesheet
            sb.Append("<!-- This is the stylesheet for the calendar -->" & Chr(13))
            sb.Append("<style TYPE=""Text/css"">" & Chr(13))
            sb.Append("<!--" & Chr(13))
            sb.Append(".CALDIM {font-family: Tahoma, Arial, Helvetica; font-size :8pt; color :#C0C0C0;text-decoration: none;}" & Chr(13))
            sb.Append(".CALGEN {font-family: Tahoma, Arial, Helvetica; font-size :8pt; color :#000000;text-decoration: none;}" & Chr(13))
            sb.Append(".CALBOLD {font-family: Tahoma, Arial, Helvetica; font-size :8pt; font-weight :bold; color :#000000;text-decoration: none;}" & Chr(13))
            sb.Append(".CALHEADER {font-family: Tahoma, Arial, Helvetica; font-size :8pt; font-weight :bold; color :" & strMonthNameColor & ";text-decoration: none;}" & Chr(13))
            sb.Append("-->" & Chr(13))
            sb.Append("</style>" & Chr(13))

            'external table that contains the calendar and form to select a different date
            sb.Append("<table WIDTH=140 BORDER=0 CELLSPACING=1 CELLPADDING=1 BGCOLOR=#FFFFFF>" & Chr(13))
            sb.Append("<tr><td align=center>")


            'CALENDAR
            sb.Append("<table WIDTH=140 BORDER=0 CELLSPACING=0 CELLPADDING=2 BGCOLOR=#666666>" & Chr(13))
            sb.Append("<tr><td>" & Chr(13))
            sb.Append("<table WIDTH=140 BORDER=0 CELLSPACING=0 CELLPADDING=1 BGCOLOR=#ffffff>" & Chr(13))
            sb.Append("<tr><td>" & Chr(13))
            sb.Append("<!--This is the table that contains the calendar-->" & Chr(13))
            sb.Append("<table id=calendar WIDTH=140 BORDER=0 CELLPADDING=1 CELLSPACING=0 BGCOLOR=#FFFFFF>" & Chr(13))

            sb.Append("<!-- Month and year name with previous and next arrows  -->" & Chr(13))

            'Calendar top menu:month name & links to previous and next month
            sb.Append("<tr WIDTH=140 BGCOLOR=" & strTopColor & ">" & Chr(13))
            sb.Append("<td WIDTH=20 HEIGHT=16 ALIGN=LEFT VALIGN=MIDDLE><a HREF=""" & strScript & "?SelectedDate=" & strPreviousDate & "&FormFieldName=" & strFormFieldName & """><img SRC=""" & strProjectPath & "images/prev_month.gif"" WIDTH=10 HEIGHT=16 BORDER=0 HSPACE=0 VSPACE=0></a></td>" & Chr(13))
            sb.Append("<td WIDTH=120 COLSPAN=5 ALIGN=CENTER VALIGN=MIDDLE CLASS=CALHEADER>" & strMonthName & " " & intThisYear & "</td>" & Chr(13))
            sb.Append("<td WIDTH=20 HEIGHT=16 ALIGN=LEFT VALIGN=MIDDLE><a HREF=""" & strScript & "?SelectedDate=" & strNextDate & "&FormFieldName=" & strFormFieldName & """><img SRC=""" & strProjectPath & "images/next_month.gif"" WIDTH=10 HEIGHT=16 BORDER=0 HSPACE=0 VSPACE=0></a></td>" & Chr(13))
            sb.Append("</tr>" & Chr(13))

            'Row with weekday name initials
            sb.Append("<tr>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Su</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Mo</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Tu</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>We</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Th</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Fr</td>" & Chr(13))
            sb.Append("<td ALIGN=RIGHT CLASS=CALGEN WIDTH=20 HEIGHT=15 VALIGN=BOTTOM>Sa</td>" & Chr(13))
            sb.Append("</tr>" & Chr(13))

            'line separating the day names and dates
            sb.Append("<tr><td HEIGHT=1 WIDTH=140 ALIGN=MIDDLE COLSPAN=7 BGCOLOR=#000000></td></tr>" & Chr(13))

            'Dates
            ' Generate 6 rows in the calendar
            For iRowNum = 0 To 5
                ' Start a table row. This corresponds to one week
                sb.Append("<TR>" & Chr(13))
                ' This is the loop for the days in the week
                For iLoopDay = 1 To 7

                    intThisDay = Day(dtMonthDay)

                    If dtMonthDay > dtLastDay Or dtMonthDay < dtFirstDay Then

                        sb.Append(Write_TD("", "CALDIM"))

                    ElseIf dtMonthDay = dtDateSelected Then

                        sb.Append(Write_TD("<a href=javascript:ReturnDate('" & dtMonthDay & "'); CLASS=CALBOLD>" & intThisDay & "</a>", "CALBOLD"))

                    Else

                        sb.Append(Write_TD("<a href=javascript:ReturnDate('" & dtMonthDay & "'); CLASS=CALGEN>" & intThisDay & "</a>", "GEN"))

                    End If

                    dtMonthDay = CDate(DateAdd(DateInterval.Day, 1, dtMonthDay))

                Next
                ' Close the table row. End of that week
                sb.Append("</TR>" & Chr(13))

            Next

            sb.Append("</table>" & Chr(13))
            sb.Append("</td></tr>" & Chr(13))
            sb.Append("</table>" & Chr(13))
            sb.Append("</td></tr>" & Chr(13))

            sb.Append("</table>" & Chr(13))

            sb.Append("</td></tr>" & Chr(13))
            'END OF CALENDAR



            'FORM
            sb.Append("<tr><td height=5></td></tr>")

            sb.Append("<tr><td align=center valign=center>")
            sb.Append(FormDate(dtDateSelected))
            sb.Append("</td></tr>")

            'end of external table
            sb.Append("</table>" & Chr(13))

            sb.Append("<script language=javascript>" & Chr(13))
            sb.Append("<!--" & Chr(13))

            sb.Append("var formElement;" & Chr(13))

            sb.Append("function ReturnDate(selDate) {" & Chr(13))
            sb.Append("window.opener.document.getElementById('" & strFormFieldName & "').value = selDate;" & Chr(13))
            sb.Append("window.close();" & Chr(13))
            sb.Append("}" & Chr(13))

            sb.Append("function changeDate() {" & Chr(13))
            sb.Append("var mDate = document.calfrm.cal_month.value + '/1/' + document.calfrm.cal_year.value;" & Chr(13))
            sb.Append("myURL = '" & strScript & "?FormFieldName=" & strFormFieldName & "&SelectedDate=' + mDate ;" & Chr(13))
            sb.Append("x = setTimeout('window.location.href=myURL',0); " & Chr(13))
            sb.Append("}" & Chr(13))

            sb.Append("//-->" & Chr(13))
            sb.Append("</script>" & Chr(13).ToString)


            Return sb.ToString

        End Function



        Function Write_TD(ByVal psValue As String, ByVal psClass As String) As String

            Return "<TD ALIGN=RIGHT WIDTH=20 HEIGHT=15 VALIGN=BOTTOM CLASS='" & psClass & "'> " & psValue & "</TD>" & Chr(13)

        End Function


        Public Function FormDate(ByVal dtPassed As Date) As String

            Dim sb As StringBuilder = New StringBuilder("")
            Dim mIndex As Integer

            Try
                sb.Append("<form action=calendar.asp method=post name=calfrm>")

                'Month
                sb.Append("<select name=""cal_month"" size=1 id=cal_month class=CALGEN onchange=changeDate();>" & Chr(13))

                For mIndex = 1 To 12

                    sb.Append("<option value=" & mIndex)

                    If mIndex = dtPassed.Month Then sb.Append(" selected ")

                    sb.Append(">" & Functions.SpanishMonthName(mIndex) & "</option>" & Chr(13))

                Next

                sb.Append("</select>" & Chr(13))


                'Year
                sb.Append("<select name=""cal_year"" size=1 id=cal_year class=CALGEN onchange=changeDate();>" & Chr(13))


                For mIndex = intYearStart To intYearEnd

                    sb.Append("<option value=" & mIndex)

                    If mIndex = dtPassed.Year Then sb.Append(" selected ")

                    sb.Append(">" & mIndex & "</option>" & Chr(13))

                Next

                sb.Append("</select>" & Chr(13))


                sb.Append("</form>")

                Return sb.ToString

            Catch e As Exception

                Return e.ToString

            Finally

                sb = Nothing

            End Try

        End Function

        
    End Class


End Namespace

