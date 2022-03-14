VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HoehleKlasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type HoehleType
    name As String
    nachbarHoehlen As ArrayList
    bewohner As BewohnerInterface
End Type
Private this As HoehleType

Private Sub Class_Initialize()

    Set this.nachbarHoehlen = New ArrayList
    Set this.bewohner = Nothing

End Sub

Public Function Constructor(ByVal neuerName As String) As HoehleKlasse
Attribute Constructor.VB_UserMemId = 0
' Default Member!
' Attribute Constructor.VB_UserMemId = 0

    Dim neueHoehle As New HoehleKlasse

    neueHoehle.name = neuerName
        
    Set Constructor = neueHoehle
        
End Function

Public Property Get name() As String
    name = this.name
End Property

Public Property Let name(ByVal newValue As String)
    this.name = newValue
End Property

Public Property Get nachbarHoehlen() As ArrayList
    nachbarHoehlen = this.nachbarHoehlen
End Property

Public Property Let nachbarHoehlen(ByVal neueNachbarhoehlen As ArrayList)
    this.nachbarHoehlen = neueNachbarhoehlen
End Property

Public Property Set nachbarHoehle(ByVal neueNachbarhoehle As HoehleKlasse)
    
    this.nachbarHoehlen.Add neueNachbarhoehle

End Property

Public Property Get bewohner()
    Set bewohner = this.bewohner
End Property

Public Property Set bewohner(ByVal neuerBewohner As BewohnerInterface)
    Set this.bewohner = neuerBewohner
End Property

Public Function IstLeer() As Boolean
    IstLeer = this.bewohner Is Nothing
End Function

