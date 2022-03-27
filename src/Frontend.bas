Attribute VB_Name = "Frontend"
'@Folder("Wumpus.Frontend")
Option Explicit

Private foo As Boolean

Public Function InputFormAbfragen(zustand As String, SpielerHoehle As String) As String
    
    Dim form As New InputForm
    
    form.ZustandLabel.Caption = zustand
    form.LandkarteImage.Picture = LoadPicture(CurrentPicture(SpielerHoehle))
    
    form.Show vbModal
    
    Dim eingabe As String
    eingabe = form.result
    Unload form
    
    InputFormAbfragen = eingabe

End Function

Private Function CurrentPicture(aktuellePosition As String) As String
    CurrentPicture = ThisWorkbook.Path & "\Landkarten\" & aktuellePosition & ".jpg"
End Function

