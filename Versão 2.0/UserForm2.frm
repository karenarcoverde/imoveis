VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Excluir Cliente"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()

    
    Range("B10").End(xlDown).Select
    linha = Cells.Find(ComboBox1.Value).Row
        
    

    nome = UserForm2.ComboBox1.Value
    
    Range("B" & linha & ":L" & linha).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    Unload UserForm2


    MsgBox ("Cliente " & nome & " excluído(a) com sucesso!")
  

End Sub


Private Sub UserForm_Initialize()

    linha = Range("B10").End(xlDown).Row
    
    ComboBox1.RowSource = "B10:B" & linha
    
    

End Sub
