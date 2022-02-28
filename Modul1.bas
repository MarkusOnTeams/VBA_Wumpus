Attribute VB_Name = "Modul1"
Option Explicit
Option Base 1

Const AnzahlHoehlen As Integer = 20

Dim Landkarte As Variant
Dim Hoehle As Variant

Dim SpielerHoehle As Integer

Sub Main()
    
    LandkarteEinlesen
    StartAufstellungFiguren_setzen
    SpielerOptionen
    
End Sub

Sub LandkarteEinlesen()

    'Landkarte = tblLandkarte.ListObjects("Verbindungen").DataBodyRange.Value
    Landkarte = tblLandkarte.Range("Verbindungen").Value
    
End Sub

Sub StartAufstellungFiguren_setzen()
    
    Hoehle = tblLandkarte.Range("Hoehle").Value
    
    Dim h As Integer
    For h = 1 To AnzahlHoehlen
        Hoehle(h, 1) = ""
    Next h
    
    SpielerHoehle = EinzelFigur_Setzen(Hoehle, "Spieler")
    
    EinzelFigur_Setzen Hoehle, "Wumpus"
    EinzelFigur_Setzen Hoehle, "Fledermaus"
    EinzelFigur_Setzen Hoehle, "Fledermaus"
    EinzelFigur_Setzen Hoehle, "Grube"
    EinzelFigur_Setzen Hoehle, "Grube"

    tblLandkarte.Range("Hoehle").Value = Hoehle


End Sub

Function EinzelFigur_Setzen(arrAufstellung As Variant, strFigur As String) As Integer
    
    Dim Zufallszahl As Integer
    
    Dim bolGefunden As Boolean
    While bolGefunden = False
        
        Zufallszahl = Int(Rnd * (AnzahlHoehlen - 1) + 1)
        
        If arrAufstellung(Zufallszahl, 1) = "" Then
            arrAufstellung(Zufallszahl, 1) = strFigur
            bolGefunden = True
            
            EinzelFigur_Setzen = Zufallszahl
        End If
    
    Wend

End Function

Sub SpielerOptionen()

Dim InHoehle As String
Dim NachHoehle1 As String
Dim NachHoehle2 As String
Dim NachHoehle3 As String

    InHoehle = Landkarte(SpielerHoehle, 1)
    NachHoehle1 = Landkarte(SpielerHoehle, 2)
    NachHoehle2 = Landkarte(SpielerHoehle, 3)
    NachHoehle3 = Landkarte(SpielerHoehle, 4)

    Debug.Print "Du bist in Höhle "; InHoehle
    Debug.Print "Es geht nach "; NachHoehle1; ", "; NachHoehle2; " und "; NachHoehle3
    
    Dim antwort As String
    
    antwort = HoehlenAntwort(NachHoehle1)
    If antwort <> "" Then Debug.Print "Es "; antwort
    
    antwort = HoehlenAntwort(NachHoehle2)
    If antwort <> "" Then Debug.Print "Es "; antwort
    
    antwort = HoehlenAntwort(NachHoehle3)
    If antwort <> "" Then Debug.Print "Es "; antwort

End Sub

Function HoehlenAntwort(HoehlenZeichen As String) As String

Dim antwort As String

Dim HoehlenNummer As Integer

    HoehlenNummer = Asc(HoehlenZeichen) - Asc("A") + 1
    Dim Bewohner As String
    
    Bewohner = Hoehle(HoehlenNummer, 1)
    Select Case Bewohner
    Case "Wumpus"
        antwort = "stinkt"
    Case "Fledermaus"
        antwort = "flattert"
    Case "Grube"
        antwort = "zieht"
    Case Else
        antwort = ""
    End Select
    
    HoehlenAntwort = antwort

End Function
