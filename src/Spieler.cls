VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Spieler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Figur")

Option Explicit

Implements FigurInterface

Private Const figurArt As String = "Spieler"
Private Const AnzahlPfeileAmAnfang As Integer = 5

Private Type figurType
    name As String
    wahrnehmung As String
    aktuelleHoehle As Hoehle
    AnzahlPfeile As Integer
End Type

Private this As figurType

Private Sub Class_Initialize()
    this.name = vbNullString
    this.wahrnehmung = vbNullString
    Set this.aktuelleHoehle = Nothing
    this.AnzahlPfeile = AnzahlPfeileAmAnfang
End Sub

Private Property Get FigurInterface_Name() As String
    FigurInterface_Name = this.name
End Property

Private Property Let FigurInterface_Name(ByVal neuerWert As String)
    this.name = neuerWert
End Property

Private Property Get FigurInterface_Art() As String
    FigurInterface_Art = figurArt
End Property

Private Property Let FigurInterface_Art(ByVal neuerWert As String)
    MsgBox "Geht nicht"
    Stop
End Property

Private Property Get FigurInterface_Wahrnehmung() As String
    FigurInterface_Wahrnehmung = this.wahrnehmung
End Property

Private Property Let FigurInterface_Wahrnehmung(ByVal neuerWert As String)
    this.wahrnehmung = neuerWert
End Property

Private Property Get FigurInterface_AktuelleHoehle() As Hoehle
    Set FigurInterface_AktuelleHoehle = this.aktuelleHoehle
End Property

Private Property Set FigurInterface_AktuelleHoehle(ByVal neueHoehle As Hoehle)
    If Not this.aktuelleHoehle Is Nothing Then Set this.aktuelleHoehle.inhalt = Nothing
    Set this.aktuelleHoehle = neueHoehle
    Set neueHoehle.inhalt = Me
End Property

Public Property Get AnzahlPfeile() As Integer
    AnzahlPfeile = this.AnzahlPfeile
End Property

Public Property Let AnzahlPfeile(ByVal neuerWert As Integer)
    this.AnzahlPfeile = neuerWert
End Property

