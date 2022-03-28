Attribute VB_Name = "SpielInfos"
Option Explicit

Private Type SpielInfosType
    
    
    verbindungen As Variant
    verbindungenDict As Scripting.Dictionary
    ' Key = Von
    ' Value = 1-basiertes Array( Nach1, ... ), nur wenn Nach(i) kein leerer String
    hoehlenAnzahl As Integer
    maxNachbarn As Integer

    hoehlen As Variant
    
    figuren As Variant
    figurenAnzahl As Integer
    FigurenDict As Scripting.Dictionary
    ' key = Figurname
    ' value = 1-basiertes Array( Figurtyp, Wahrnehmung )

End Type

Private this As SpielInfosType

Public Property Get count() As Integer

    count = IIf(Not this.verbindungenDict Is Nothing, this.verbindungenDict.count, 0)

End Property

Public Property Get HoehlenNamen() As String()

    Dim namen() As String
    ReDim namen(1 To this.verbindungenDict.count)
    
    Dim Index As Integer
    Index = 1
    
    Dim name As Variant
    For Each name In this.verbindungenDict.Keys
        namen(Index) = name
        Index = Index + 1
    Next name
    
    HoehlenNamen = namen
    
End Property


Public Property Get hoehlennachbarNamen(ByVal hoehlenName As String) As String()
    
    hoehlennachbarNamen = this.verbindungenDict.Item(hoehlenName)

End Property

Public Sub Initialisieren()

    this.verbindungen = Empty
    Set this.verbindungenDict = New Scripting.Dictionary
    this.hoehlenAnzahl = 0
    this.maxNachbarn = 0
    
    this.hoehlen = Empty
    
    this.figuren = Empty
    this.figurenAnzahl = 0
    Set this.FigurenDict = New Scripting.Dictionary
    
    Einlesen

End Sub

Private Sub Einlesen()

    verbindungenLesen
    verbindungenDictFuellen
    
    hoehlenLesen

    figurenLesen
    figurenDictFuellen
    
End Sub

Private Sub verbindungenLesen()

    this.verbindungen = SpielInfosWs.Range("Verbindungen").Value
    
    this.hoehlenAnzahl = UBound(this.verbindungen, 1)
    this.maxNachbarn = UBound(this.verbindungen, 2)
    
End Sub
Private Sub verbindungenSchreiben()

    SpielInfosWs.Range("Verbindungen").Value = this.verbindungen
    
End Sub

Private Sub verbindungenDictFuellen()
    
    Dim Index As Integer
    For Index = 1 To this.hoehlenAnzahl
    
        Dim hoehlenName As String
        hoehlenName = this.verbindungen(Index, 1)

    this.verbindungenDict.Add hoehlenName, nachbarNamen(Index)
        
    Next Index
    
End Sub

Private Sub hoehlenLesen()

    this.hoehlen = SpielInfosWs.Range("Hoehlen").Value

End Sub

Private Sub hoehlenSchreiben()

    SpielInfosWs.Range("Hoehlen").Value = this.hoehlen

End Sub

Private Sub figurenLesen()

    this.figuren = SpielInfosWs.Range("Figuren").Value

    this.figurenAnzahl = UBound(this.figuren, 1)

End Sub

Private Sub figurenSchreiben()

    SpielInfosWs.Range("Figuren").Value = this.figuren

End Sub

Private Sub figurenDictFuellen()
    
    Dim Index As Integer
    For Index = 1 To this.figurenAnzahl
    
        Dim figurName As String
        figurName = this.figuren(Index, 1)
        
        Dim eigenschaften(1 To 2) As String
        eigenschaften(1) = this.figuren(Index, 2)
        eigenschaften(2) = this.figuren(Index, 3)

    this.FigurenDict.Add figurName, eigenschaften
        
    Next Index
    
End Sub

Private Function nachbarNamen(ByVal hoehlenindex As Integer) As String()

    Dim namen() As String
    Dim namenColl As New Collection
    
    Dim Index As Integer
    For Index = 2 To this.maxNachbarn
        
        Dim nachbarName As Variant
        nachbarName = this.verbindungen(hoehlenindex, Index)
        
        If nachbarName <> vbNullString Then namenColl.Add nachbarName
    
    Next Index
    
    ReDim namen(1 To namenColl.count)
    Index = 1
    For Each nachbarName In namenColl
        namen(Index) = nachbarName
        Index = Index + 1
    Next nachbarName
    
    nachbarNamen = namen

End Function

Public Property Let hoehlenInhalte(ByRef inhalte() As String)

    Dim Index As Integer
    For Index = 1 To UBound(this.hoehlen, 1)
        this.hoehlen(Index, 1) = inhalte(Index)
    Next Index

    hoehlenSchreiben

End Property

Public Property Get FigurenDict() As Scripting.Dictionary

    Set FigurenDict = this.FigurenDict

End Property
