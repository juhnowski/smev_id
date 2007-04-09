VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LCID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
'Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Sub LCID_init()

    Dim sys_LCID
    sys_LCID = GetSystemDefaultLCID()
    
    Dim usr_LCID
    usr_LCID = GetUserDefaultLCID()
    
    If sys_LCID <> usr_LCID Then
        lcid = sys_LCID
    Else
        lcid = usr_LCID
    End If
End Sub

Public Property Get lcid() As Variant
    lcid = mLocaleID
End Property

Public Property Let lcid(ByVal vNewValue As Variant)
    mLocaleID = vNewValue
End Property


