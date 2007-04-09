VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UdostDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pVid As String
Private pSeriya As String
Private pNomer As String
Private pFio As String
Private pPol As String
Private pData_rozhdeniya As String
Private pData_vidachi As String
Private pKod_podrazdeleniya As String
Private pMesto_rozhdeniya As String

Property Get mesto_rozhdeniya() As String
    mesto_rozhdeniya = pMesto_rozhdeniya
End Property
Property Let mesto_rozhdeniya(s As String)
    pMesto_rozhdeniya = s
End Property

Property Get kod_podrazdeleniya() As String
    kod_podrazdeleniya = pKod_podrazdeleniya
End Property
Property Let kod_podrazdeleniya(s As String)
    pKod_podrazdeleniya = s
End Property

Property Get data_vidachi() As String
    data_vidachi = pData_vidachi
End Property
Property Let data_vidachi(s As String)
    pData_vidachi = s
End Property

Property Get data_rozhdeniya() As String
    data_rozhdeniya = pData_rozhdeniya
End Property
Property Let data_rozhdeniya(s As String)
    pData_rozhdeniya = s
End Property
Property Get pol() As String
    pol = pPol
End Property
Property Let pol(s As String)
    pPol = s
End Property
Property Get fio() As String
    fio = pFio
End Property

Property Let fio(s As String)
    pFio = s
End Property
Property Get nomer() As String
    nomer = pNomer
End Property

Property Let nomer(s As String)
    pNomer = s
End Property

Property Get seriya() As String
    seriya = pSeriya
End Property

Property Let seriya(s As String)
    pSeriya = s
End Property
Property Get vid() As String
    vid = pVid
End Property

Property Let vid(s As String)
    pVid = s
End Property

