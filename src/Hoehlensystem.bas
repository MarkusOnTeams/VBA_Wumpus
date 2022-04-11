Attribute VB_Name = "Hoehlensystem"
Option Explicit

Private Type HoehlensystemType
    HoehlenNamen As Variant
    HoehlenDict As Scripting.Dictionary
End Type

Private this As HoehlensystemType

Public Sub Aufbauen()

    Set this.HoehlenDict = New Scripting.Dictionary
    
    SpielInfos.Initialisieren
    
    HoehlenOhneNachbarnAufbauen
    
    HoehlenMitIhrenNachbarnVerbinden

End Sub
Private Sub HoehlenOhneNachbarnAufbauen()

    this.HoehlenNamen = SpielInfos.HoehlenNamen
    
    Dim Index As Integer
    For Index = 1 To UBound(this.HoehlenNamen)
        
        Dim hoehlenName As String
        hoehlenName = this.HoehlenNamen(Index)
        
        Dim neueHoehle As hoehle
        'Set neueHoehle = Hoehle.Constructor(hoehlenName)
        Set neueHoehle = hoehle(hoehlenName)
        
        this.HoehlenDict.Add hoehlenName, neueHoehle
    
    Next Index
    
End Sub
Private Sub HoehlenMitIhrenNachbarnVerbinden()

    
    Dim Index As Integer
    For Index = 1 To UBound(this.HoehlenNamen)
    
        Dim hoehlenName As String
        hoehlenName = this.HoehlenNamen(Index)
        
        Dim hoehle As hoehle
        Set hoehle = this.HoehlenDict(hoehlenName)
    
        Dim nachbarNamen As Variant
        nachbarNamen = SpielInfos.hoehlennachbarNamen(hoehlenName)
                    
        Dim nachbarIndex As Integer
        For nachbarIndex = 1 To UBound(nachbarNamen)
                    
            Dim nachbarName As Variant
            nachbarName = nachbarNamen(nachbarIndex)
                
            Dim nachbarHoehle As hoehle
            Set nachbarHoehle = this.HoehlenDict(nachbarName)
            
            Set hoehle.nachbarHoehle = nachbarHoehle
        
        Next nachbarIndex
        
    Next Index

End Sub

Public Function FreieHoehle() As hoehle

    Dim aktuelleHoehle As hoehle

    Dim zufallszahl As Integer
    
    Dim bolGefunden As Boolean
    Do While bolGefunden = False
        
        zufallszahl = Int(Rnd * UBound(this.HoehlenNamen) + 1)
        
        Dim hoehlenName As String
        hoehlenName = this.HoehlenNamen(zufallszahl)
        
        Set aktuelleHoehle = this.HoehlenDict(hoehlenName)
        
        bolGefunden = aktuelleHoehle.IstLeer
        If bolGefunden Then Exit Do
    
    Loop

    If bolGefunden Then Set FreieHoehle = aktuelleHoehle

End Function

Public Function hoehlenInhalte() As String()
    
    Dim figuren() As String
    ReDim figuren(UBound(this.HoehlenNamen))
    
    Dim Index As Integer
    For Index = 1 To UBound(this.HoehlenNamen)
    
        Dim hoehlenName As String
        hoehlenName = this.HoehlenNamen(Index)
    
        Dim hoehle As hoehle
        Set hoehle = this.HoehlenDict(hoehlenName)
    
        Dim figurName As String
        figurName = vbNullString
        If Not hoehle.IstLeer Then figurName = hoehle.inhalt.name
        
        figuren(Index) = figurName
    
    Next Index
    
    hoehlenInhalte = figuren
    
End Function

Public Property Get HoehlenNamen()
    HoehlenNamen = this.HoehlenNamen
End Property

Public Property Get Index(ByVal name As String)
    
    Index = 0

    Dim i As Integer
    For i = 1 To this.HoehlenNamen
        If this.HoehlenNamen(i) = name Then
            
            Index = i
            Exit For
        
        End If
    Next i

End Property

Public Property Get Item(ByVal name As String)
    Set Item = this.HoehlenDict(name)
End Property
