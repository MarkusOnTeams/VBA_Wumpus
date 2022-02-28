Attribute VB_Name = "Modul1"
Option Explicit
Option Base 1

Const AnzahlHoehlen As Integer = 20
Const AnzahlPfeileAmAnfang As Integer = 5

Const AktionGehen As String = "_"
Const AktionSchiessen As String = ">"

Const Spieler As String = "Spieler"
Const Fledermaus As String = "Fledermaus"
Const Wumpus As String = "Wumpus"
Const Grube As String = "Grube"

Dim Landkarte As Variant
Dim hoehleninhalt As Variant

Dim SpielerHoehlenNummer As Integer
Dim WumpusHoehlenNummer As Integer

Dim WumpusInNachbarhoehle As Boolean

Dim SpielerHoehle As String
Dim NachHoehle1 As String
Dim NachHoehle2 As String
Dim NachHoehle3 As String

Dim MoeglicheAktionen As Scripting.Dictionary

Dim Navi(1 To 4) As Variant

Dim AnzahlPfeile As Integer

Dim SpielLaeuft As Boolean

Sub Main()
    
    LandkarteEinlesen
    StartAufstellungFiguren_setzen
    
    AnzahlPfeile = AnzahlPfeileAmAnfang
    
    SpielLaeuft = True
    Do While SpielLaeuft
        
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

Sub LandkarteEinlesen()

    'Landkarte = LandkarteWs.ListObjects("Verbindungen").DataBodyRange.Value
    Landkarte = LandkarteWs.Range("Verbindungen").Value
    
End Sub

Sub StartAufstellungFiguren_setzen()
    
    hoehleninhalt = LandkarteWs.Range("Hoehle").Value
    
    Dim h As Integer
    For h = 1 To AnzahlHoehlen
        hoehleninhalt(h, 1) = ""
    Next h
    
    SpielerHoehlenNummer = EinzelFigur_Setzen(hoehleninhalt, Spieler)
    
    WumpusHoehlenNummer = EinzelFigur_Setzen(hoehleninhalt, Wumpus)
    
    EinzelFigur_Setzen hoehleninhalt, Fledermaus
    EinzelFigur_Setzen hoehleninhalt, Fledermaus
    EinzelFigur_Setzen hoehleninhalt, Grube
    EinzelFigur_Setzen hoehleninhalt, Grube

    LandkarteWs.Range("Hoehle").Value = hoehleninhalt


End Sub

Function EinzelFigur_Setzen(arrAufstellung As Variant, strFigur As String) As Integer
    
    Dim zufallszahl As Integer
    
    Dim bolGefunden As Boolean
    While bolGefunden = False
        
        zufallszahl = Int(Rnd * (AnzahlHoehlen - 1) + 1)
        
        If arrAufstellung(zufallszahl, 1) = "" Then
            arrAufstellung(zufallszahl, 1) = strFigur
            bolGefunden = True
            
            EinzelFigur_Setzen = zufallszahl
        End If
    
    Wend

End Function

Function SpielZustand() As String

Dim ergebnis As String

    SpielerHoehle = Landkarte(SpielerHoehlenNummer, 1)
    NachHoehle1 = Landkarte(SpielerHoehlenNummer, 2)
    NachHoehle2 = Landkarte(SpielerHoehlenNummer, 3)
    NachHoehle3 = Landkarte(SpielerHoehlenNummer, 4)
    
    Navi(1) = NachHoehle1
    Navi(2) = NachHoehle2
    Navi(3) = NachHoehle3
    Navi(4) = SpielerHoehle

    ergebnis = "Du bist in Höhle " & SpielerHoehle & vbLf
    ergebnis = ergebnis & "Es geht nach " & NachHoehle1 & ", " & NachHoehle2 & " und " & NachHoehle3 & vbLf
    
    WumpusInNachbarhoehle = False
    
    Dim wahrnehmung As String
    wahrnehmung = HoehlenAntwort(NachHoehle1)
    If wahrnehmung <> "" Then ergebnis = ergebnis & "Es " & wahrnehmung & vbLf
    
    wahrnehmung = HoehlenAntwort(NachHoehle2)
    If wahrnehmung <> "" Then ergebnis = ergebnis & "Es " & wahrnehmung & vbLf
    
    wahrnehmung = HoehlenAntwort(NachHoehle3)
    If wahrnehmung <> "" Then ergebnis = ergebnis & "Es " & wahrnehmung & vbLf
    
    ergebnis = ergebnis & "Du hast noch " & AnzahlPfeile & " Pfeil" & IIf(AnzahlPfeile > 1, "e", "") & vbLf
    
    
    If MoeglicheAktionen Is Nothing Then Set MoeglicheAktionen = New Scripting.Dictionary
    MoeglicheAktionen.RemoveAll
    
    Dim aktion As String
    ergebnis = ergebnis & "Deine möglichen Aktionen: "
    Dim trenner As String
    trenner = ""
    Dim i As Integer
    For i = 1 To 3
        aktion = AktionGehen & Navi(i)
        MoeglicheAktionen.Add aktion, Null
        ergebnis = ergebnis & trenner & aktion
        trenner = "  "
    Next i
    For i = 1 To 3
        aktion = AktionSchiessen & Navi(i)
        MoeglicheAktionen.Add aktion, Null
        ergebnis = ergebnis & trenner & aktion
        trenner = "  "
    Next i
    
    SpielZustand = ergebnis

End Function

Function HoehlenAntwort(HoehlenZeichen As String) As String

Dim antwort As String

Dim hoehlennummer As Integer

    hoehlennummer = Asc(HoehlenZeichen) - Asc("A") + 1
    Dim bewohner As String
    
    bewohner = hoehleninhalt(hoehlennummer, 1)
    Select Case bewohner
    Case Wumpus
        antwort = "stinkt"
        WumpusInNachbarhoehle = True
    Case Fledermaus
        antwort = "flattert"
    Case Grube
        antwort = "zieht"
    Case Else
        antwort = ""
    End Select
    
    HoehlenAntwort = antwort

End Function

Function SpielerAktion(zustand As String) As String

    Dim zulaessigeAntwort As Boolean
    zulaessigeAntwort = False
    
    Do While Not zulaessigeAntwort
    
        Dim eingabe As String
        eingabe = InputBox(prompt:=zustand, Title:="Deine Aktion?")
        
        If eingabe = vbNullString Then
        
            zulaessigeAntwort = True
        
        Else
            
            zulaessigeAntwort = MoeglicheAktionen.Exists(eingabe)
            
            If Not zulaessigeAntwort Then MsgBox "Deine Antwort " & eingabe & " ist nicht zulässig. Nochmal ..."
            
        End If

    Loop
    
    SpielerAktion = eingabe

End Function

Sub NeuerSpielzustand(aktion As String)

    Dim gehenOderSchiessen As String
    Dim hoehle As String
    
    gehenOderSchiessen = Left(aktion, 1)
    hoehle = Mid(aktion, 2, 1)
    
    Select Case gehenOderSchiessen
    
    Case AktionGehen
        GeheNachHoehle hoehle
    
    Case AktionSchiessen
        SchiesseInHoehle hoehle
    
    End Select

End Sub

Sub GeheNachHoehle(hoehle As String)

    Dim hoehlennummer As Integer
    hoehlennummer = Asc(hoehle) - Asc("A") + 1
    
    Dim bewohner As String
    bewohner = hoehleninhalt(hoehlennummer, 1)
    
    Select Case bewohner
    Case Wumpus
        MsgBox "Der Wumpus hat dich gefressen"
        SpielLaeuft = False
    
    Case Fledermaus
        hoehleninhalt(SpielerHoehlenNummer, 1) = ""
        hoehlennummer = EinzelFigur_Setzen(hoehleninhalt, Spieler)
        SpielerHoehlenNummer = hoehlennummer
        hoehleninhalt(SpielerHoehlenNummer, 1) = Spieler
    
    Case Grube
        MsgBox "Du bist in eine bodenlose Grube gestürzt"
        SpielLaeuft = False
    
    Case Else
        hoehleninhalt(SpielerHoehlenNummer, 1) = ""
        SpielerHoehlenNummer = hoehlennummer
        hoehleninhalt(SpielerHoehlenNummer, 1) = Spieler
    
    End Select

    LandkarteWs.Range("Hoehle").Value = hoehleninhalt

End Sub

Sub SchiesseInHoehle(hoehle As String)

    Dim hoehlennummer As Integer
    hoehlennummer = Asc(hoehle) - Asc("A") + 1
    
    Dim bewohner As String
    bewohner = hoehleninhalt(hoehlennummer, 1)
    
    Select Case bewohner
    Case Wumpus
    
        MsgBox "Du hast den Wumpus getötet. Herzlichen Glückwunsch!"
        SpielLaeuft = False
    
    Case Else
        If WumpusInNachbarhoehle And WumpusWachtAuf() Then WumpusBewegtSich
    
    End Select
    
    AnzahlPfeile = AnzahlPfeile - 1
    If SpielLaeuft And AnzahlPfeile = 0 Then
        MsgBox "Du hast keine Pfeile mehr. Dumm gelaufen ..."
        SpielLaeuft = False
    End If
    
    LandkarteWs.Range("Hoehle").Value = hoehleninhalt

End Sub

Function WumpusWachtAuf() As Boolean

    WumpusWachtAuf = Int(Rnd() * 100) <= 75

End Function

Sub WumpusBewegtSich()

    Dim nachbarHoehlen As Variant

    nachbarHoehlen = Array(Landkarte(WumpusHoehlenNummer, 2), Landkarte(WumpusHoehlenNummer, 3), Landkarte(WumpusHoehlenNummer, 4))
    
    Dim nachbarHoehlenNummer(3)
    
    Dim i As Integer
    For i = 1 To 3
        nachbarHoehlenNummer(i) = Asc(nachbarHoehlen(i)) - Asc("A") + 1
        
        Dim bewohner As String
        bewohner = hoehleninhalt(nachbarHoehlenNummer(i), 1)
        
        If bewohner = Fledermaus Or bewohner = Grube Then nachbarHoehlenNummer(i) = 0
        
    Next i
    
    Dim zufallszahl As Integer
    zufallszahl = Int(Rnd() * 3 + 1)
    
    Do While nachbarHoehlenNummer(zufallszahl) = 0
        zufallszahl = zufallszahl + 1
        If zufallszahl > 3 Then zufallszahl = 1
    Loop
    
    Dim nachHoehlennummer As Integer
    nachHoehlennummer = nachbarHoehlenNummer(zufallszahl)
    
    bewohner = hoehleninhalt(nachHoehlennummer, 1)
    If bewohner = Spieler Then
        MsgBox "Der Wumpus ist aufgewacht, hat dich gefunden und gefressen..."
        SpielLaeuft = False
    Else
        hoehleninhalt(WumpusHoehlenNummer, 1) = ""
        WumpusHoehlenNummer = nachHoehlennummer
        hoehleninhalt(WumpusHoehlenNummer, 1) = Wumpus
        MsgBox "Der Wumpus ist aufgewacht und in eine andere Höhle gewandert"
    End If
    
End Sub
