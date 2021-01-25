VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_QuantityOrdered As Integer
Private m_QuantityNested As Integer
Private m_Priority As Integer
Private m_Material As String
Private m_Thickness As Double
Private m_CpFile As String
Private m_Turret As String


Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal sName As String)
    m_Name = VBA.Left(sName, 70)
End Property

Public Property Get QuantityOrdered() As Integer
    QuantityOrdered = m_QuantityOrdered
End Property

Public Property Let QuantityOrdered(ByVal iQuantityOrdered As Integer)
    m_QuantityOrdered = iQuantityOrdered
End Property

Public Property Get QuantityNested() As Integer
    QuantityNested = m_QuantityNested
End Property

Public Property Let QuantityNested(ByVal iQuantityNested As Integer)
    m_QuantityNested = iQuantityNested
End Property

Public Property Get Priority() As Integer
    Priority = m_Priority
End Property

Public Property Let Priority(ByVal iPriority As Integer)
    m_Priority = iPriority
End Property

Public Property Get Status() As Integer
    Status = IIf(m_QuantityOrdered - m_QuantityNested <= 0, 2, 0)
End Property

Public Property Get Material() As String
    Material = m_Material
End Property

Public Property Let Material(ByVal sMaterial As String)
    m_Material = sMaterial
End Property

Public Property Get Thickness() As Double
    Thickness = m_Thickness
End Property

Public Property Let Thickness(ByVal iThickness As Double)
    m_Thickness = iThickness
End Property

Public Property Get CpFile() As String
    CpFile = m_CpFile
End Property

Public Property Let CpFile(ByVal sCpFile As String)
    m_CpFile = sCpFile
End Property

Public Property Get Turret() As String
    Turret = m_Turret
End Property

Public Property Let Turret(ByVal sTurret As String)
    m_Turret = sTurret
End Property
