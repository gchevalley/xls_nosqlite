VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_NoSQL_ObjectId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'a voir plutot comme un generator
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Private timestamp_unix As Double
'Private machine_id As Double
'Private process_id As Double
'Private counter As Double

    Private last_timestamp As Long
    Private objectid_count As Long



Private Sub Class_Initialize()

last_timestamp = Me.get_current_timestamp_unix()
objectid_count = 0

End Sub


'Public Function init_with_json(ByVal str_json As String)
'
'End Function



'Public Function get_new_format_json() As String
'
'Dim oJSON As New jsonlib
'
'Debug.Print get_current_timestamp_unix()
'
'
'End Function


Public Function get_current_timestamp_unix() As Long

Dim date_start_unix As Date
    date_start_unix = DateSerial(1970, 1, 1)
    
get_current_timestamp_unix = DateDiff("s", date_start_unix, Now)

End Function


Public Function get_next_json() As String

Dim oJSON As New jsonlib
get_next_json = oJSON.toString(Me.get_next_dic())

End Function


Public Function get_next_dic() As Scripting.Dictionary

Dim tmp_timestamp_unix As Long
    tmp_timestamp_unix = Me.get_current_timestamp_unix
    
    If tmp_timestamp_unix <> last_timestamp Then
        last_timestamp = tmp_timestamp_unix
    End If

Dim lpBuffC As String * 1314
GetComputerName lpBuffC, Len(lpBuffC)

Dim lpBuffU As String * 25
GetUserName lpBuffU, 25

Randomize


Dim dic_oid As Scripting.Dictionary
Set dic_oid = New Scripting.Dictionary
    dic_oid.Add "timestamp_unix", tmp_timestamp_unix
    dic_oid.Add "milliseconds", CInt(Right(Format(Timer, "#0.00"), 2))
    dic_oid.Add "machine_id", Application.Trim(Application.Clean(lpBuffC))
    dic_oid.Add "username", Left(lpBuffU, InStr(lpBuffU, Chr(0)) - 1)
    dic_oid.Add "process_id", GetCurrentProcessId()
    dic_oid.Add "random", Round(1000000000 * Rnd(), 0)
    dic_oid.Add "counter", objectid_count
    



'Dim oJSON As New jsonlib
'Debug.Print oJSON.toString(dic_oid)
Set get_next_dic = dic_oid

increase_counter

End Function


Private Function increase_counter()

objectid_count = objectid_count + 1

End Function
