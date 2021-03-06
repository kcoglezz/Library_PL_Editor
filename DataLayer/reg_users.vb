Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer

    Public Class reg_users

        '*********************************************************************
        '  Class Name: reg_users
        '  Author: Automated Code Generator
        '  Date Created: 4/13/2009 10:02:33 AM
        '  Revisions:
        '
        '*********************************************************************

#Region "Constructor"

        Public Sub New()
        End Sub

        Public Sub New(ByVal intId As Integer)


            If intId <> 0 Then

                getByID(intId)

            End If

        End Sub

#End Region

#Region "Variables"

        Private _intId As Integer = 0

        Private _strFirstName As String = ""

        Private _strLastName As String = ""

        Private _strPassword As String = ""

        Private _strEmail As String = ""

        Private _btFirstVisit As Byte = 0

        Private _dtBday As Date = #1/1/1900#

        Private _btSex As Byte = 0

        Private _btMStatus As Byte = 0

        Private _btWorkType As Byte = 0

        Private _strCity As String = ""

        Private _strState As String = ""

        Private _btCountry As Byte = 0

        Private _strMainLanguage As String = ""

        Private _strPhone As String = ""

        Private _intSecurityLevel As Integer = 0

        Private _strPicture As String = ""

        Private _strOtherinfo As String = ""

        Private _dtLastLog As Date = #1/1/1900#

        Private _btShowInfo As Byte = 0

        Private _btReceiveEmails As Byte = 0

#End Region

#Region "Properties"

        'Properties For Id
        Public Property Id() As Integer

            Get

                Return _intId

            End Get

            Set(ByVal Value As Integer)

                _intId = Value

            End Set

        End Property

        Public Sub setId(ByVal Value As Integer)

            _intId = Value

        End Sub
        Public Function getId() As Integer

            Return _intId

        End Function

        'Properties For FirstName
        Public Property FirstName() As String

            Get

                Return _strFirstName

            End Get

            Set(ByVal Value As String)

                _strFirstName = Value

            End Set

        End Property

        Public Sub setFirstName(ByVal Value As String)

            _strFirstName = Value

        End Sub
        Public Function getFirstName() As String

            Return _strFirstName

        End Function

        'Properties For LastName
        Public Property LastName() As String

            Get

                Return _strLastName

            End Get

            Set(ByVal Value As String)

                _strLastName = Value

            End Set

        End Property

        Public Sub setLastName(ByVal Value As String)

            _strLastName = Value

        End Sub
        Public Function getLastName() As String

            Return _strLastName

        End Function

        'Properties For Password
        Public Property Password() As String

            Get

                Return _strPassword

            End Get

            Set(ByVal Value As String)

                _strPassword = Value

            End Set

        End Property

        Public Sub setPassword(ByVal Value As String)

            _strPassword = Value

        End Sub
        Public Function getPassword() As String

            Return _strPassword

        End Function

        'Properties For Email
        Public Property Email() As String

            Get

                Return _strEmail

            End Get

            Set(ByVal Value As String)

                _strEmail = Value

            End Set

        End Property

        Public Sub setEmail(ByVal Value As String)

            _strEmail = Value

        End Sub
        Public Function getEmail() As String

            Return _strEmail

        End Function

        'Properties For FirstVisit
        Public Property FirstVisit() As Byte

            Get

                Return _btFirstVisit

            End Get

            Set(ByVal Value As Byte)

                _btFirstVisit = Value

            End Set

        End Property

        Public Sub setFirstVisit(ByVal Value As Byte)

            _btFirstVisit = Value

        End Sub
        Public Function getFirstVisit() As Byte

            Return _btFirstVisit

        End Function

        'Properties For Bday
        Public Property Bday() As Date

            Get

                Return _dtBday

            End Get

            Set(ByVal Value As Date)

                _dtBday = Value

            End Set

        End Property

        Public Sub setBday(ByVal Value As Date)

            _dtBday = Value

        End Sub
        Public Function getBday() As Date

            Return _dtBday

        End Function

        'Properties For Sex
        Public Property Sex() As Byte

            Get

                Return _btSex

            End Get

            Set(ByVal Value As Byte)

                _btSex = Value

            End Set

        End Property

        Public Sub setSex(ByVal Value As Byte)

            _btSex = Value

        End Sub
        Public Function getSex() As Byte

            Return _btSex

        End Function

        'Properties For MStatus
        Public Property MStatus() As Byte

            Get

                Return _btMStatus

            End Get

            Set(ByVal Value As Byte)

                _btMStatus = Value

            End Set

        End Property

        Public Sub setMStatus(ByVal Value As Byte)

            _btMStatus = Value

        End Sub
        Public Function getMStatus() As Byte

            Return _btMStatus

        End Function

        'Properties For WorkType
        Public Property WorkType() As Byte

            Get

                Return _btWorkType

            End Get

            Set(ByVal Value As Byte)

                _btWorkType = Value

            End Set

        End Property

        Public Sub setWorkType(ByVal Value As Byte)

            _btWorkType = Value

        End Sub
        Public Function getWorkType() As Byte

            Return _btWorkType

        End Function

        'Properties For City
        Public Property City() As String

            Get

                Return _strCity

            End Get

            Set(ByVal Value As String)

                _strCity = Value

            End Set

        End Property

        Public Sub setCity(ByVal Value As String)

            _strCity = Value

        End Sub
        Public Function getCity() As String

            Return _strCity

        End Function

        'Properties For State
        Public Property State() As String

            Get

                Return _strState

            End Get

            Set(ByVal Value As String)

                _strState = Value

            End Set

        End Property

        Public Sub setState(ByVal Value As String)

            _strState = Value

        End Sub
        Public Function getState() As String

            Return _strState

        End Function

        'Properties For Country
        Public Property Country() As Byte

            Get

                Return _btCountry

            End Get

            Set(ByVal Value As Byte)

                _btCountry = Value

            End Set

        End Property

        Public Sub setCountry(ByVal Value As Byte)

            _btCountry = Value

        End Sub
        Public Function getCountry() As Byte

            Return _btCountry

        End Function

        'Properties For MainLanguage
        Public Property MainLanguage() As String

            Get

                Return _strMainLanguage

            End Get

            Set(ByVal Value As String)

                _strMainLanguage = Value

            End Set

        End Property

        Public Sub setMainLanguage(ByVal Value As String)

            _strMainLanguage = Value

        End Sub
        Public Function getMainLanguage() As String

            Return _strMainLanguage

        End Function

        'Properties For Phone
        Public Property Phone() As String

            Get

                Return _strPhone

            End Get

            Set(ByVal Value As String)

                _strPhone = Value

            End Set

        End Property

        Public Sub setPhone(ByVal Value As String)

            _strPhone = Value

        End Sub
        Public Function getPhone() As String

            Return _strPhone

        End Function

        'Properties For SecurityLevel
        Public Property SecurityLevel() As Integer

            Get

                Return _intSecurityLevel

            End Get

            Set(ByVal Value As Integer)

                _intSecurityLevel = Value

            End Set

        End Property

        Public Sub setSecurityLevel(ByVal Value As Integer)

            _intSecurityLevel = Value

        End Sub
        Public Function getSecurityLevel() As Integer

            Return _intSecurityLevel

        End Function

        'Properties For Picture
        Public Property Picture() As String

            Get

                Return _strPicture

            End Get

            Set(ByVal Value As String)

                _strPicture = Value

            End Set

        End Property

        Public Sub setPicture(ByVal Value As String)

            _strPicture = Value

        End Sub
        Public Function getPicture() As String

            Return _strPicture

        End Function

        'Properties For Otherinfo
        Public Property Otherinfo() As String

            Get

                Return _strOtherinfo

            End Get

            Set(ByVal Value As String)

                _strOtherinfo = Value

            End Set

        End Property

        Public Sub setOtherinfo(ByVal Value As String)

            _strOtherinfo = Value

        End Sub
        Public Function getOtherinfo() As String

            Return _strOtherinfo

        End Function

        'Properties For LastLog
        Public Property LastLog() As Date

            Get

                Return _dtLastLog

            End Get

            Set(ByVal Value As Date)

                _dtLastLog = Value

            End Set

        End Property

        Public Sub setLastLog(ByVal Value As Date)

            _dtLastLog = Value

        End Sub
        Public Function getLastLog() As Date

            Return _dtLastLog

        End Function

        'Properties For ShowInfo
        Public Property ShowInfo() As Byte

            Get

                Return _btShowInfo

            End Get

            Set(ByVal Value As Byte)

                _btShowInfo = Value

            End Set

        End Property

        Public Sub setShowInfo(ByVal Value As Byte)

            _btShowInfo = Value

        End Sub
        Public Function getShowInfo() As Byte

            Return _btShowInfo

        End Function

        'Properties For ReceiveEmails
        Public Property ReceiveEmails() As Byte

            Get

                Return _btReceiveEmails

            End Get

            Set(ByVal Value As Byte)

                _btReceiveEmails = Value

            End Set

        End Property

        Public Sub setReceiveEmails(ByVal Value As Byte)

            _btReceiveEmails = Value

        End Sub
        Public Function getReceiveEmails() As Byte

            Return _btReceiveEmails

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbRegUsers" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetRegUsersByID " & intId)

                    setValues(reader)

                    HttpContext.Current.Trace.Write("Data Layer Object retrieved database")

                    HttpContext.Current.Cache.Insert(strCacheVar, Me)

                Catch ex As Exception

                    Throw ex

                Finally

                    reader.Close()
                    reader = Nothing
                End Try

            Else 'If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim objCached As Object = HttpContext.Current.Cache.Get(strCacheVar)

                Try

                    HttpContext.Current.Trace.Write("Data Layer Object retrieved from cache")

                    _intId = objCached.getId

                    _strFirstName = objCached.getFirstName

                    _strLastName = objCached.getLastName

                    _strPassword = objCached.getPassword

                    _strEmail = objCached.getEmail

                    _btFirstVisit = objCached.getFirstVisit

                    _dtBday = objCached.getBday

                    _btSex = objCached.getSex

                    _btMStatus = objCached.getMStatus

                    _btWorkType = objCached.getWorkType

                    _strCity = objCached.getCity

                    _strState = objCached.getState

                    _btCountry = objCached.getCountry

                    _strMainLanguage = objCached.getMainLanguage

                    _strPhone = objCached.getPhone

                    _intSecurityLevel = objCached.getSecurityLevel

                    _strPicture = objCached.getPicture

                    _strOtherinfo = objCached.getOtherinfo

                    _dtLastLog = objCached.getLastLog

                    _btShowInfo = objCached.getShowInfo

                    _btReceiveEmails = objCached.getReceiveEmails

                Catch ex As Exception

                    Throw ex

                Finally

                    objCached = Nothing

                End Try

            End If 'If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

        End Sub


        'Sub setValues:  Sets the values to all properties for this object from database

        Private Sub setValues(ByVal reader As System.Data.SqlClient.SqlDataReader)
            Try
                If reader.HasRows() Then

                    reader.Read()

                    If Not reader.IsDBNull(0) Then

                        _intId = reader(0)

                    End If


                    If Not reader.IsDBNull(1) Then

                        _strFirstName = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _strLastName = reader(2)

                    End If


                    If Not reader.IsDBNull(3) Then

                        _strPassword = reader(3)

                    End If


                    If Not reader.IsDBNull(4) Then

                        _strEmail = reader(4)

                    End If


                    If Not reader.IsDBNull(5) Then

                        _btFirstVisit = reader(5)

                    End If


                    If Not reader.IsDBNull(6) Then

                        _dtBday = reader(6)

                    End If


                    If Not reader.IsDBNull(7) Then

                        _btSex = reader(7)

                    End If


                    If Not reader.IsDBNull(8) Then

                        _btMStatus = reader(8)

                    End If


                    If Not reader.IsDBNull(9) Then

                        _btWorkType = reader(9)

                    End If


                    If Not reader.IsDBNull(10) Then

                        _strCity = reader(10)

                    End If


                    If Not reader.IsDBNull(11) Then

                        _strState = reader(11)

                    End If


                    If Not reader.IsDBNull(12) Then

                        _btCountry = reader(12)

                    End If


                    If Not reader.IsDBNull(13) Then

                        _strMainLanguage = reader(13)

                    End If


                    If Not reader.IsDBNull(14) Then

                        _strPhone = reader(14)

                    End If


                    If Not reader.IsDBNull(15) Then

                        _intSecurityLevel = reader(15)

                    End If


                    If Not reader.IsDBNull(16) Then

                        _strPicture = reader(16)

                    End If


                    If Not reader.IsDBNull(17) Then

                        _strOtherinfo = reader(17)

                    End If


                    If Not reader.IsDBNull(18) Then

                        _dtLastLog = reader(18)

                    End If


                    If Not reader.IsDBNull(19) Then

                        _btShowInfo = reader(19)

                    End If


                    If Not reader.IsDBNull(20) Then

                        _btReceiveEmails = reader(20)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbRegUsers" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteRegUsers " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strFirstName As String, ByVal strLastName As String, ByVal strPassword As String, ByVal strEmail As String, ByVal btFirstVisit As Byte, ByVal dtBday As Date, ByVal btSex As Byte, ByVal btMStatus As Byte, ByVal btWorkType As Byte, ByVal strCity As String, ByVal strState As String, ByVal btCountry As Byte, ByVal strMainLanguage As String, ByVal strPhone As String, ByVal intSecurityLevel As Integer, ByVal strPicture As String, ByVal strOtherinfo As String, ByVal dtLastLog As Date, ByVal btShowInfo As Byte, ByVal btReceiveEmails As Byte)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertRegUsers '" & strFirstName & "','" & strLastName & "','" & strPassword & "','" & strEmail & "'," & btFirstVisit & ",'" & dtBday & "'," & btSex & "," & btMStatus & "," & btWorkType & ",'" & strCity & "','" & strState & "'," & btCountry & ",'" & strMainLanguage & "','" & strPhone & "'," & intSecurityLevel & ",'" & strPicture & "','" & strOtherinfo & "','" & dtLastLog & "'," & btShowInfo & "," & btReceiveEmails))

                _intId = intId
                _strFirstName = strFirstName
                _strLastName = strLastName
                _strPassword = strPassword
                _strEmail = strEmail
                _btFirstVisit = btFirstVisit
                _dtBday = dtBday
                _btSex = btSex
                _btMStatus = btMStatus
                _btWorkType = btWorkType
                _strCity = strCity
                _strState = strState
                _btCountry = btCountry
                _strMainLanguage = strMainLanguage
                _strPhone = strPhone
                _intSecurityLevel = intSecurityLevel
                _strPicture = strPicture
                _strOtherinfo = strOtherinfo
                _dtLastLog = dtLastLog
                _btShowInfo = btShowInfo
                _btReceiveEmails = btReceiveEmails

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertRegUsers '" & _strFirstName & "','" & _strLastName & "','" & _strPassword & "','" & _strEmail & "'," & _btFirstVisit & ",'" & _dtBday & "'," & _btSex & "," & _btMStatus & "," & _btWorkType & ",'" & _strCity & "','" & _strState & "'," & _btCountry & ",'" & _strMainLanguage & "','" & _strPhone & "'," & _intSecurityLevel & ",'" & _strPicture & "','" & _strOtherinfo & "','" & _dtLastLog & "'," & _btShowInfo & "," & _btReceiveEmails))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal strFirstName As String, ByVal strLastName As String, ByVal strPassword As String, ByVal strEmail As String, ByVal btFirstVisit As Byte, ByVal dtBday As Date, ByVal btSex As Byte, ByVal btMStatus As Byte, ByVal btWorkType As Byte, ByVal strCity As String, ByVal strState As String, ByVal btCountry As Byte, ByVal strMainLanguage As String, ByVal strPhone As String, ByVal intSecurityLevel As Integer, ByVal strPicture As String, ByVal strOtherinfo As String, ByVal dtLastLog As Date, ByVal btShowInfo As Byte, ByVal btReceiveEmails As Byte)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateRegUsers " & intId & ",'" & strFirstName & "','" & strLastName & "','" & strPassword & "','" & strEmail & "'," & btFirstVisit & ",'" & dtBday & "'," & btSex & "," & btMStatus & "," & btWorkType & ",'" & strCity & "','" & strState & "'," & btCountry & ",'" & strMainLanguage & "','" & strPhone & "'," & intSecurityLevel & ",'" & strPicture & "','" & strOtherinfo & "','" & dtLastLog & "'," & btShowInfo & "," & btReceiveEmails)

                _intId = intId
                _strFirstName = strFirstName
                _strLastName = strLastName
                _strPassword = strPassword
                _strEmail = strEmail
                _btFirstVisit = btFirstVisit
                _dtBday = dtBday
                _btSex = btSex
                _btMStatus = btMStatus
                _btWorkType = btWorkType
                _strCity = strCity
                _strState = strState
                _btCountry = btCountry
                _strMainLanguage = strMainLanguage
                _strPhone = strPhone
                _intSecurityLevel = intSecurityLevel
                _strPicture = strPicture
                _strOtherinfo = strOtherinfo
                _dtLastLog = dtLastLog
                _btShowInfo = btShowInfo
                _btReceiveEmails = btReceiveEmails

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateRegUsers " & _intId & ",'" & _strFirstName & "','" & _strLastName & "','" & _strPassword & "','" & _strEmail & "'," & _btFirstVisit & ",'" & _dtBday & "'," & _btSex & "," & _btMStatus & "," & _btWorkType & ",'" & _strCity & "','" & _strState & "'," & _btCountry & ",'" & _strMainLanguage & "','" & _strPhone & "'," & _intSecurityLevel & ",'" & _strPicture & "','" & _strOtherinfo & "','" & _dtLastLog & "'," & _btShowInfo & "," & _btReceiveEmails)


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.reg_users

            Dim strCacheVar As String = "objCachedDbRegUsers" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.reg_users

                objCached.setId(_intId)

                objCached.setFirstName(_strFirstName)

                objCached.setLastName(_strLastName)

                objCached.setPassword(_strPassword)

                objCached.setEmail(_strEmail)

                objCached.setFirstVisit(_btFirstVisit)

                objCached.setBday(_dtBday)

                objCached.setSex(_btSex)

                objCached.setMStatus(_btMStatus)

                objCached.setWorkType(_btWorkType)

                objCached.setCity(_strCity)

                objCached.setState(_strState)

                objCached.setCountry(_btCountry)

                objCached.setMainLanguage(_strMainLanguage)

                objCached.setPhone(_strPhone)

                objCached.setSecurityLevel(_intSecurityLevel)

                objCached.setPicture(_strPicture)

                objCached.setOtherinfo(_strOtherinfo)

                objCached.setLastLog(_dtLastLog)

                objCached.setShowInfo(_btShowInfo)

                objCached.setReceiveEmails(_btReceiveEmails)

                HttpContext.Current.Cache.Insert(strCacheVar, objCached)

            Catch ex As Exception

                Throw ex

            Finally

                objCached = Nothing

            End Try

        End Sub




#End Region


    End Class

End Namespace
