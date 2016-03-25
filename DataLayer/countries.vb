Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



    Public Class countries

        '*********************************************************************
        '  Class Name: countries
        '  Author: Automated Code Generator
        '  Date Created: 3/18/2009 3:28:49 PM
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

        Private _strName As String = ""

        Private _strCountryCode As String = ""

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

        'Properties For Name
        Public Property Name() As String

            Get

                Return _strName

            End Get

            Set(ByVal Value As String)

                _strName = Value

            End Set

        End Property

        Public Sub setName(ByVal Value As String)

            _strName = Value

        End Sub
        Public Function getName() As String

            Return _strName

        End Function

        'Properties For CountryCode
        Public Property CountryCode() As String

            Get

                Return _strCountryCode

            End Get

            Set(ByVal Value As String)

                _strCountryCode = Value

            End Set

        End Property

        Public Sub setCountryCode(ByVal Value As String)

            _strCountryCode = Value

        End Sub
        Public Function getCountryCode() As String

            Return _strCountryCode

        End Function
#End Region

#Region "Methods"

        'Sub getByID:  Gets all the values for this record
        Public Sub getByID(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbCountries" & intId

            If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                Dim reader As System.Data.SqlClient.SqlDataReader

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetCountriesByID " & intId)

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

                    _strName = objCached.getName

                    _strCountryCode = objCached.getCountryCode

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

                        _strName = reader(1)

                    End If


                    If Not reader.IsDBNull(2) Then

                        _strCountryCode = reader(2)

                    End If



                End If 'If reader.HasRows())

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Delete function

        Public Sub Delete(ByVal intId As Integer)

            Dim strCacheVar As String = "objCachedDbCountries" & intId

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteCountries " & intId)

                If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

                    HttpContext.Current.Cache.Remove(strCacheVar)

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add(ByVal strName As String, ByVal strCountryCode As String)

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertCountries '" & strName & "','" & strCountryCode & "'"))

                _intId = intId
                _strName = strName
                _strCountryCode = strCountryCode

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Add Sub

        Public Sub Add()

            Dim intId As Integer = 0

            Try

                intId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertCountries '" & _strName & "','" & _strCountryCode & "'"))

                _intId = intId

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub
        'Update Sub

        Public Sub Update(ByVal intId As Integer, ByVal strName As String, ByVal strCountryCode As String)

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateCountries " & intId & ",'" & strName & "','" & strCountryCode & "'")

                _intId = intId
                _strName = strName
                _strCountryCode = strCountryCode

                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        'Update Sub

        Public Sub Update()

            Try

                ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateCountries " & _intId & ",'" & _strName & "','" & _strCountryCode & "'")


                UpdateCache()


            Catch ex As Exception

                Throw ex

            End Try

        End Sub

        Public Sub UpdateCache()

            Dim objCached As DataLayer.countries

            Dim strCacheVar As String = "objCachedDbCountries" & _intId

            Try

                HttpContext.Current.Cache.Remove(strCacheVar)

                objCached = New DataLayer.countries

                objCached.setId(_intId)

                objCached.setName(_strName)

                objCached.setCountryCode(_strCountryCode)

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
