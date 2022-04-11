VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wumpus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Figur")

Option Explicit

Implements FigurInterface

Private Const figurArt As String = "Wumpus"

Private Type figurType
    name As String
    wahrnehmung As String
    aktuelleHoehle As Hoehle
End Type

Private this As figurType

Private Sub Class_Initialize()
    this.name = vbNullString
    this.wahrnehmung = vbNullString
    Set this.aktuelleHoehle = Nothing
End Sub

Private Property Get FigurInterface_name() As String
    FigurInterface_name = this.name
End Property

Private Property Let FigurInterface_name(ByVal neuerWert As String)
    this.name = neuerWert
End Property

Private Property Get FigurInterface_Art() As String
    FigurInterface_Art = figurArt
End Property

Private Property Get FigurInterface_wahrnehmung() As String
    FigurInterface_wahrnehmung = this.wahrnehmung
End Property

Private Property Let FigurInterface_wahrnehmung(ByVal neuerWert As String)
    this.wahrnehmung = neuerWert
End Property

Private Property Get FigurInterface_aktuelleHoehle() As Hoehle
    Set FigurInterface_aktuelleHoehle = this.aktuelleHoehle
End Property

Private Property Set FigurInterface_aktuelleHoehle(ByVal neueHoehle As Hoehle)
    If Not this.aktuelleHoehle Is Nothing Then Set this.aktuelleHoehle.inhalt = Nothing
    Set this.aktuelleHoehle = neueHoehle
End Property

Private Sub FigurInterface_Kollision()
    MsgBox "Der Wumpus hat dich gefressen"
    SpielLaeuft = False
End Sub
