Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Public FigurenDict As Scripting.Dictionary
Public SpielerFigur As FigurInterface
Public SpielerAlsSpieler As spieler
Public WumpusFigur As FigurInterface

Public WumpusInNachbarhoehle As Boolean

Const AktionGehen As String = "_"
Const AktionSchiessen As String = ">"

Public MoeglicheAktionen As Scripting.Dictionary

Public SpielLaeuft As Boolean

Sub Main()
    
    Hoehlensystem.Aufbauen
    
    Hoehlen_fuellen Figuren_generieren()
    
    SpielInfos.hoehlenInhalte = Hoehlensystem.hoehlenInhalte
    
    SpielLaeuft = True
    Do While SpielLaeuft
    
        DoEvents
        
        Dim zustand As String
        zustand = SpielZustand()

        Dim aktion As String
        aktion = SpielerAktion(zustand)
        Debug.Print aktion
        If aktion = vbNullString Then
            SpielLaeuft = False
        End If

        NeuerSpielzustand aktion
        
    Loop
    
End Sub

Private Function Figuren_generieren() As Collection

    Dim figuren As New Collection

    Dim FigurenDict As Scripting.Dictionary
    Set FigurenDict = SpielInfos.FigurenDict
    
    Dim figurName As Variant
    For Each figurName In FigurenDict.Keys
    
        Dim eigenschaften As Variant
        eigenschaften = FigurenDict(figurName)
        
        Dim figurArt As String
        figurArt = eigenschaften(1)
        
        Dim wahrnehmung As String
        wahrnehmung = eigenschaften(2)
        
        Dim neueFigur As FigurInterface
        
        Select Case figurArt
        Case "Spieler"
            Set neueFigur = New spieler
            Set SpielerFigur = neueFigur
            Set SpielerAlsSpieler = neueFigur
        
        Case "Wumpus"
            Set neueFigur = New Wumpus
            Set WumpusFigur = neueFigur
        
        Case "Fledermaus"
            Set neueFigur = New Fledermaus
        
        Case "Grube"
            Set neueFigur = New Grube
        
        Case Else
        
            MsgBox "unbekannte Art " & figurArt
            Stop
        
        End Select

        neueFigur.name = figurName
        neueFigur.wahrnehmung = wahrnehmung
        
        figuren.Add neueFigur

    Next figurName
    
    Set Figuren_generieren = figuren

End Function

Private Sub Hoehlen_fuellen(ByVal figuren As Collection)

    Dim figur As FigurInterface
    For Each figur In figuren
        
        Dim neueHoehle As Hoehle
        Set neueHoehle = Hoehlensystem.FreieHoehle
        
        Set figur.aktuelleHoehle = neueHoehle
        Set neueHoehle.inhalt = figur
        'Debug.Print neueHoehle.name, figur.name, figur.wahrnehmung
        
    Next figur

End Sub

Function SpielZustand() As String

    Dim ergebnis As String

    Dim spielerHoehle As String
    spielerHoehle = SpielerFigur.aktuelleHoehle.name
    
    Dim nachbarHoehlen() As Hoehle
    nachbarHoehlen = SpielerFigur.aktuelleHoehle.nachbarHoehlen
    
    Dim anzahlNachbarhoehlen As Integer
    anzahlNachbarhoehlen = UBound(nachbarHoehlen)
    
    Dim nachHoehle() As String
    nachHoehle = SpielerFigur.aktuelleHoehle.nachbarHoehlenNamen
    
    ergebnis = "Du bist in Höhle " & spielerHoehle & vbLf
    ergebnis = ergebnis & "Es geht nach "
    
    Dim trenner As String
    
    Dim Index As Integer
    For Index = 1 To anzahlNachbarhoehlen
        ergebnis = ergebnis & trenner & nachHoehle(Index)
        trenner = ", "
    Next Index
    ergebnis = ergebnis & vbLf
    
    Dim wahrnehmung As String
    wahrnehmung = ""
    For Index = 1 To anzahlNachbarhoehlen
        If Not nachbarHoehlen(Index).IstLeer Then
            wahrnehmung = nachbarHoehlen(Index).inhalt.wahrnehmung
            ergebnis = ergebnis & "Es " & wahrnehmung & vbLf
        End If
    Next Index

    WumpusInNachbarhoehle = InStr(wahrnehmung, WumpusFigur.wahrnehmung)

    ergebnis = ergebnis & "Du hast noch " & SpielerAlsSpieler.AnzahlPfeile & " Pfeil" & IIf(SpielerAlsSpieler.AnzahlPfeile > 1, "e", "") & vbLf

    If MoeglicheAktionen Is Nothing Then Set MoeglicheAktionen = New Scripting.Dictionary
    MoeglicheAktionen.RemoveAll

    Dim aktion As String
    ergebnis = ergebnis & "Deine möglichen Aktionen: "
    trenner = ""
    Dim i As Integer
    For i = 1 To anzahlNachbarhoehlen
        aktion = AktionGehen & nachHoehle(i)
        MoeglicheAktionen.Add aktion, Null
        ergebnis = ergebnis & trenner & aktion
        trenner = "  "
    Next i
    For i = 1 To anzahlNachbarhoehlen
        aktion = AktionSchiessen & nachHoehle(i)
        MoeglicheAktionen.Add aktion, Null
        ergebnis = ergebnis & trenner & aktion
        trenner = "  "
    Next i

    SpielZustand = ergebnis

End Function

Function SpielerAktion(zustand As String) As String

    Dim zulaessigeAntwort As Boolean
    zulaessigeAntwort = False

    Do While Not zulaessigeAntwort

        Dim eingabe As String
        ' eingabe = InputBox(prompt:=zustand, Title:="Deine Aktion?")

        eingabe = InputFormAbfragen(zustand)

        If eingabe = vbNullString Then

            zulaessigeAntwort = True

        Else

            zulaessigeAntwort = MoeglicheAktionen.Exists(eingabe)

            If Not zulaessigeAntwort Then MsgBox "Deine Antwort " & eingabe & " ist nicht zulässig. Nochmal ..."

        End If

    Loop

    SpielerAktion = eingabe

End Function

Function InputFormAbfragen(zustand As String) As String

    InputForm.ZustandLabel.Caption = zustand
    InputForm.EingabeTextbox.Text = ""

    Dim Bildname As String
    Bildname = ThisWorkbook.Path & "\Landkarten\" & SpielerFigur.aktuelleHoehle.name & ".jpg"
    InputForm.LandkarteImage.Picture = LoadPicture(Bildname)

    InputForm.Show vbModal

    ' InputForm.Hide

    Dim eingabe As String
    eingabe = InputForm.EingabeTextbox.Text

    InputFormAbfragen = eingabe

End Function

Sub NeuerSpielzustand(aktion As String)

    Dim gehenOderSchiessen As String
    Dim Hoehle As String

    gehenOderSchiessen = Left(aktion, 1)
    Hoehle = Mid(aktion, 2)

    Select Case gehenOderSchiessen

    Case AktionGehen
        GeheNachHoehle Hoehle

    Case AktionSchiessen
        SchiesseInHoehle Hoehle

    End Select

End Sub

Sub GeheNachHoehle(hoehlenName As String)

    Dim nachHoehle As Hoehle
    Set nachHoehle = Hoehlensystem.Item(hoehlenName)

    If nachHoehle.IstLeer Then
        Set SpielerFigur.aktuelleHoehle = nachHoehle
    Else

'        Dim figur As FigurInterface
'        Set figur = nachHoehle.inhalt
'        figur.Kollision
        
        nachHoehle.inhalt.Kollision
        
'        Select Case figur.Art
'        Case "Wumpus"
'            MsgBox "Der Wumpus hat dich gefressen"
'            SpielLaeuft = False
'
'        Case "Fledermaus"
'            MsgBox "Die Fledermaus " & figur.name & " hat dich in eine andere Höhle verschleppt"
'            Set SpielerFigur.aktuelleHoehle = Hoehlensystem.FreieHoehle
'
'        Case "Grube"
'            MsgBox "Du bist in eine bodenlose Grube gestürzt"
'            SpielLaeuft = False
'
'        Case Else
'            MsgBox "Unbekannte Art " & figur.Art
'            Stop
'        End Select

    End If
    
    SpielInfos.hoehlenInhalte = Hoehlensystem.hoehlenInhalte

End Sub

Sub SchiesseInHoehle(ByVal hoehlenName As String)

    Dim nachHoehle As Hoehle
    Set nachHoehle = Hoehlensystem.Item(hoehlenName)

    Dim figur As FigurInterface
    Set figur = nachHoehle.inhalt

    Select Case figur.Art
        Case "Wumpus"

            MsgBox "Du hast den Wumpus getötet. Herzlichen Glückwunsch!"
            SpielLaeuft = False

        Case Else
            If WumpusInNachbarhoehle And WumpusWachtAuf() Then WumpusBewegtSich

    End Select

    SpielerAlsSpieler.AnzahlPfeile = SpielerAlsSpieler.AnzahlPfeile - 1
    If SpielLaeuft And SpielerAlsSpieler.AnzahlPfeile = 0 Then
        MsgBox "Du hast keine Pfeile mehr. Dumm gelaufen ..."
        SpielLaeuft = False
    End If

    SpielInfos.hoehlenInhalte = Hoehlensystem.hoehlenInhalte

End Sub

Function WumpusWachtAuf() As Boolean

    WumpusWachtAuf = Int(Rnd() * 100) <= 75

End Function

Sub WumpusBewegtSich()

    Dim nachbarHoehlen() As Hoehle
    nachbarHoehlen = WumpusFigur.aktuelleHoehle.nachbarHoehlen

    Dim zufallszahl As Integer
    zufallszahl = Int(Rnd() * UBound(nachbarHoehlen) + 1)

    Dim nachHoehle As Hoehle
    Do
        Set nachHoehle = nachbarHoehlen(zufallszahl)
        If nachHoehle.IstLeer Then Exit Do
            
        Dim figur As FigurInterface
        Set figur = nachHoehle.inhalt
        
        If figur.Art = "Spieler" Then
            MsgBox "Der Wumpus ist aufgewacht, hat dich gefunden und gefressen..."
            SpielLaeuft = False
            Exit Do
        End If
        
        zufallszahl = zufallszahl + 1
        If zufallszahl > 3 Then zufallszahl = 1
    Loop
            
    Set WumpusFigur.aktuelleHoehle = nachHoehle
    MsgBox "Der Wumpus ist aufgewacht und in eine andere Höhle gewandert"

End Sub


