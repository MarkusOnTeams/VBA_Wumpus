VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoehle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type HoehleType
    name As String
    nachbarHoehlen() As Hoehle
    figur As FigurInterface
End Type
Private this As HoehleType

Private Sub Class_Initialize()

    Set this.figur = Nothing

End Sub

Public Function Constructor(ByVal neuerName As String) As Hoehle
Attribute Constructor.VB_UserMemId = 0
' Default Member!
' Attribute Constructor.VB_UserMemId = 0

    Dim neueHoehle As New Hoehle

    neueHoehle.name = neuerName
        
    Set Constructor = neueHoehle
        
End Function

Public Property Get name() As String
    name = this.name
End Property

Public Property Let name(ByVal newValue As String)
    this.name = newValue
End Property

Public Property Get nachbarHoehlen() As Hoehle()
    nachbarHoehlen = this.nachbarHoehlen
End Property

Public Property Let nachbarHoehlen(ByRef neueNachbarhoehlen() As Hoehle)
    this.nachbarHoehlen = neueNachbarhoehlen
End Property

Public Property Get nachbarHoehlenNamen() As String()

    Dim count As Integer
    
    count = UBound(this.nachbarHoehlen)
    
    Dim namen() As String
    ReDim namen(1 To count)
    
    Dim Index As Integer
    For Index = 1 To count
        
        namen(Index) = this.nachbarHoehlen(Index).name

    Next Index
    
    nachbarHoehlenNamen = namen

End Property

Public Property Set nachbarHoehle(ByVal neueNachbarhoehle As Hoehle)
    
    Dim count As Integer
    
    count = 0
    If Not Not this.nachbarHoehlen Then count = UBound(this.nachbarHoehlen)
    
    ReDim Preserve this.nachbarHoehlen(1 To count + 1)
    Set this.nachbarHoehlen(count + 1) = neueNachbarhoehle
    
End Property

Public Property Get inhalt()
    Set inhalt = this.figur
End Property

Public Property Set inhalt(ByVal neueFigur As FigurInterface)
    Set this.figur = neueFigur
End Property

Public Function IstLeer() As Boolean
    IstLeer = this.figur Is Nothing
End Function

