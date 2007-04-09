VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransSredstva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActualnost As String
Private pKategoriya As String
Private pMarka As String
Private pModel As String
Private pCvet As String
Private pGosZnak As String
Private pVIN As String
Private pNDvig As String
Private pKodPodr As String
Private pGodVipuska As String

Property Get GodVipuska() As String
    GodVipuska = pGodVipuska
End Property

Property Let GodVipuska(s As String)
    pGodVipuska = s
End Property

Property Get KodPodr() As String
    KodPodr = pKodPodr
End Property

Property Let KodPodr(s As String)
    pKodPodr = s
End Property

Property Get NDvig() As String
    NDvig = pNDvig
End Property

Property Let NDvig(s As String)
    pNDvig = s
End Property


Property Get VIN() As String
    VIN = pVIN
End Property

Property Let VIN(s As String)
    pVIN = s
End Property

Property Get GosZnak() As String
    GosZnak = pGosZnak
End Property

Property Let GosZnak(s As String)
    pGosZnak = s
End Property

Property Get Cvet() As String
    Cvet = pCvet
End Property

Property Let Cvet(s As String)
    pCvet = s
End Property

Property Get Model() As String
    Model = pModel
End Property

Property Let Model(s As String)
    pModel = s
End Property

Property Get Marka() As String
    Marka = pMarka
End Property

Property Let Marka(s As String)
    pMarka = s
End Property

Property Get Kategoriya() As String
    Kategoriya = pKategoriya
End Property

Property Let Kategoriya(s As String)
    pKategoriya = s
End Property

Property Get Actualnost() As String
    Actualnost = pActualnost
End Property

Property Let Actualnost(s As String)
    pActualnost = s
End Property

