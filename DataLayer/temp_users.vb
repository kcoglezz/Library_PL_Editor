Imports System
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Web

Namespace DataLayer



Public Class temp_users  

'*********************************************************************
'  Class Name: temp_users
'  Author: Automated Code Generator
'  Date Created: 3/18/2009 3:28:50 PM
'  Revisions:
'
'*********************************************************************

 #region "Constructor" 

Public Sub New()
End Sub

Public Sub New (ByVal intColId As Integer)


If intColId <> 0 Then

getByID(intColId)

End If

End Sub

#End Region 

#region "Variables" 

Private _intColId As Integer = 0

Private _intRegUserId As Integer = 0

#End Region

#region "Properties" 

'Properties For ColId
Public Property ColId() As Integer

Get 

Return _intColId

End Get

Set(ByVal Value As  Integer)

_intColId = Value

End Set

End Property

Public Sub setColId(ByVal Value As Integer)

_intColId = Value

End Sub
Public Function getColId() As Integer

Return _intColId

End Function

'Properties For RegUserId
Public Property RegUserId() As Integer

Get 

Return _intRegUserId

End Get

Set(ByVal Value As  Integer)

_intRegUserId = Value

End Set

End Property

Public Sub setRegUserId(ByVal Value As Integer)

_intRegUserId = Value

End Sub
Public Function getRegUserId() As Integer

Return _intRegUserId

End Function
#End Region

#region "Methods"

'Sub getByID:  Gets all the values for this record
Public Sub getByID(ByVal intColId As Integer)

Dim strCacheVar As String = "objCachedDbTempUsers" & intColId

If IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then 

Dim reader As System.Data.SqlClient.SqlDataReader

Try 

reader = ParaLideres.GenericDataHandler.GetRecords("proc_GetTempUsersByID "&intColId) 

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

_intColId = objCached.getColId

_intRegUserId = objCached.getRegUserId

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

_intColId = reader(0)

End If


If Not reader.IsDBNull(1) Then 

_intRegUserId = reader(1)

End If



End If 'If reader.HasRows())

Catch ex As Exception

Throw ex 

End Try 

End Sub

'Delete function

Public Sub Delete(ByVal intColId As Integer)

Dim strCacheVar As String = "objCachedDbTempUsers" &  intColId

Try

ParaLideres.GenericDataHandler.ExecSQL("proc_DeleteTempUsers "&intColId)

If Not IsNothing(HttpContext.Current.Cache.Get(strCacheVar)) Then

HttpContext.Current.Cache.Remove(strCacheVar)

End If

Catch ex As Exception

Throw ex

End Try

End Sub

'Add Sub

Public Sub Add(ByVal intRegUserId As Integer)

Dim intColId As Integer = 0

Try

intColId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertTempUsers " & intRegUserId))

_intColId = intColId
_intRegUserId = intRegUserId

UpdateCache()


Catch ex As Exception

Throw ex

End Try

End Sub

'Add Sub

Public Sub Add()

Dim intColId As Integer = 0

Try

intColId = CInt(ParaLideres.GenericDataHandler.ExecScalar("proc_InsertTempUsers " & _intRegUserId))

_intColId = intColId

UpdateCache()


Catch ex As Exception

Throw ex

End Try

End Sub
'Update Sub

Public Sub Update(ByVal intColId As Integer,ByVal intRegUserId As Integer)

Try

ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateTempUsers " & intColId & "," & intRegUserId)

_intColId = intColId
_intRegUserId = intRegUserId

UpdateCache()


Catch ex As Exception

Throw ex

End Try

End Sub

'Update Sub

Public Sub Update()

Try

ParaLideres.GenericDataHandler.ExecSQL("proc_UpdateTempUsers " & _intColId & "," & _intRegUserId)


UpdateCache()


Catch ex As Exception

Throw ex

End Try

End Sub

Public Sub UpdateCache()

Dim objCached As DataLayer.temp_users

Dim strCacheVar As String = "objCachedDbTempUsers" &  _intColId

Try

HttpContext.Current.Cache.Remove(strCacheVar)

objCached = New DataLayer.temp_users

objCached.setColId(_intColId)

objCached.setRegUserId(_intRegUserId)

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
