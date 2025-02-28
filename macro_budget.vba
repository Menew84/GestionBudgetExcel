Option Explicit

' Macro pour insérer des cases à cocher (ActiveX) dans la colonne D (Payé ?) pour les Revenus
Sub AjouterCheckboxRevenus()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' Assurez-vous d'être sur la feuille mensuelle voulue
    
    Dim i As Long
    Dim chk As OLEObject
    
    For i = 6 To 55
        ' Supprimer la case existante si elle existe déjà
        On Error Resume Next
        ws.OLEObjects("ChkRev" & i).Delete
        On Error GoTo 0
        
        ' Ajouter une case à cocher dans la colonne D (colonne 4)
        Set chk = ws.OLEObjects.Add(ClassType:="Forms.CheckBox.1", _
                    Left:=ws.Cells(i, 4).Left + 2, Top:=ws.Cells(i, 4).Top + 2, _
                    Width:=ws.Cells(i, 4).Width - 4, Height:=ws.Cells(i, 4).Height - 4)
        chk.Name = "ChkRev" & i
        chk.Object.Caption = ""  ' Pas de texte
    Next i
End Sub

' Macro pour insérer des cases à cocher (ActiveX) dans la colonne I (Payé ?) pour les Dépenses
Sub AjouterCheckboxDepenses()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' Assurez-vous d'être sur la feuille mensuelle voulue
    
    Dim i As Long
    Dim chk As OLEObject
    
    For i = 6 To 55
        ' Supprimer la case existante si elle existe déjà
        On Error Resume Next
        ws.OLEObjects("ChkDep" & i).Delete
        On Error GoTo 0
        
        ' Ajouter une case à cocher dans la colonne I (colonne 9)
        Set chk = ws.OLEObjects.Add(ClassType:="Forms.CheckBox.1", _
                    Left:=ws.Cells(i, 9).Left + 2, Top:=ws.Cells(i, 9).Top + 2, _
                    Width:=ws.Cells(i, 9).Width - 4, Height:=ws.Cells(i, 9).Height - 4)
        chk.Name = "ChkDep" & i
        chk.Object.Caption = ""
    Next i
End Sub
