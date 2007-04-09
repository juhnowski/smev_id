VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nedvizhimost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActualnost As String
Private pNaimenovanie As String
Private pPloshad As String
Private pUslNomer As String
Private pInvNomer As String
Private pKadastrNomer As String
Private pTochAdres As String

Property Get TochAdres() As String
    TochAdres = pTochAdres
End Property

Property Let TochAdres(s As String)
    pTochAdres = s
End Property

Property Get KadastrNomer() As String
    KadastrNomer = pKadastrNomer
End Property

Property Let KadastrNomer(s As String)
    pKadastrNomer = s
End Property

Property Get InvNomer() As String
    InvNomer = pInvNomer
End Property

Property Let InvNomer(s As String)
    pInvNomer = s
End Property

Property Get UslNomer() As String
    UslNomer = pUslNomer
End Property

Property Let UslNomer(s As String)
    pUslNomer = s
End Property

Property Get Ploshad() As String
    Ploshad = pPloshad
End Property

Property Let Ploshad(s As String)
    pPloshad = s
End Property

Property Get Naimenovanie() As String
    Naimenovanie = pNaimenovanie
End Property

Property Let Naimenovanie(s As String)
    pNaimenovanie = s
End Property

Property Get Actualnost() As String
    Actualnost = pActualnost
End Property

Property Let Actualnost(s As String)
    pActualnost = s
End Property

