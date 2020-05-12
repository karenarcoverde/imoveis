VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registro de Cliente"
   ClientHeight    =   6696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7500
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OptionButton6_Click()
    UserForm1.TextBox8.Visible = OptionButton6.Value
End Sub
Private Sub OptionButton1_Click()
    If OptionButton6.Value = False Then
        UserForm1.TextBox8.Visible = False
    End If
End Sub

Private Sub OptionButton4_Click()
    If OptionButton6.Value = False Then
        UserForm1.TextBox8.Visible = False
    End If
End Sub

Private Sub OptionButton5_Click()
    If OptionButton6.Value = False Then
        UserForm1.TextBox8.Visible = False
    End If
End Sub



Private Sub OptionButton2_Click()
    UserForm1.Frame1.Visible = OptionButton2.Value
End Sub
Private Sub OptionButton3_Click()
    If OptionButton2.Value = False Then
        UserForm1.Frame1.Visible = False
    End If
End Sub



Private Sub OptionButton7_Click()
    If OptionButton8.Value = False Then
        UserForm1.Frame3.Visible = False
        UserForm1.Frame2.Visible = OptionButton7.Value
    End If
    
End Sub
Private Sub OptionButton8_Click()
    If OptionButton7.Value = False Then
        UserForm1.Frame2.Visible = False
        UserForm1.Frame3.Visible = OptionButton8.Value
    End If
    
End Sub



Private Sub TextBox2_change()
    If Len(TextBox2.Text) = 2 Then
        TextBox2 = TextBox2 + "/"
    End If
    
    If Len(TextBox2.Text) = 5 Then
        TextBox2.Text = TextBox2 + "/"
    End If

End Sub

Private Sub TextBox6_change()
    If Len(TextBox6.Text) = 2 Then
        TextBox6 = TextBox6 + "/"
    End If
    
    If Len(TextBox6.Text) = 5 Then
        TextBox6.Text = TextBox6 + "/"
    End If

End Sub




Private Sub TextBox7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 33 Then
        TextBox7.Text = TextBox7.Text & Chr(33)
    End If

End Sub




Private Sub TextBox7_change()
    Dim qtd As Integer
    
    qtd = Len(TextBox7.Text)
    
    Label8.Caption = "Quantidade de Caracteres: " & qtd

End Sub

Private Sub ToggleButton1_Click()
    If Len(TextBox7.Text) > 66 Then
        MsgBox "Ultrapassa 66 Caracteres!", vbCritical, "HISTÓRICO"
        TextBox7.SelStart = 0
        TextBox7.SelLength = TextBox7.TextLength
    End If
    
     Dim objeto As Control

    Call Registrar

    UserForm1.Hide
    
    For Each objeto In UserForm1.Controls
        On Error Resume Next
        objeto.Value = ""
    Next

End Sub



Sub Registrar()
    
    Dim range1 As Range
    
    If Range("B10").Value = "" Then
        Set range1 = Range("B10")
    Else
        Set range1 = Range("B9").End(xlDown).Offset(1, 0)
    End If
    
    
    
    range1.Value = UserForm1.TextBox1.Value
    range1.Offset(0, 1).Value = CDate(UserForm1.TextBox2.Value)
    range1.Offset(0, 2).Value = UserForm1.TextBox3.Value
    range1.Offset(0, 3).Value = UserForm1.TextBox4.Value
    
    
    'página Origem
    If UserForm1.OptionButton1.Value = True Then
        range1.Offset(0, 4).Value = "Facebook"
    ElseIf UserForm1.OptionButton4.Value = True Then
        range1.Offset(0, 4).Value = "Zap"
    ElseIf UserForm1.OptionButton5.Value = True Then
        range1.Offset(0, 4).Value = "OLX"
    ElseIf UserForm1.OptionButton6.Value = True Then
        range1.Offset(0, 4).Value = TextBox8.Value
    End If
    
    
    'página Visita
    If UserForm1.OptionButton2.Value = True Then
        range1.Offset(0, 5).Value = CDate(TextBox6.Value)
    ElseIf UserForm1.OptionButton3.Value = True Then
        range1.Offset(0, 5).Value = ""
    End If
    
    'página Imóvel
    If UserForm1.OptionButton7.Value = True Then
        range1.Offset(0, 6).Value = "Lançamento"
        range1.Offset(0, 7).Value = "Nome do Empreendimento: " & TextBox9.Value
    ElseIf UserForm1.OptionButton8.Value = True Then
        range1.Offset(0, 6).Value = "Usado"
        range1.Offset(0, 7).Value = TextBox10.Value & " - " & TextBox11.Value
    End If
    
    
    
    'página Tipo de Cliente
    If UserForm1.OptionButton9.Value = True Then
        range1.Offset(0, 8).Value = "Potencial"
    ElseIf UserForm1.OptionButton10.Value = True Then
        range1.Offset(0, 8).Value = "Pesquisando"
    ElseIf UserForm1.OptionButton13.Value = True Then
        range1.Offset(0, 8).Value = "Frio"
    End If
    
    'página Histórico
    range1.Offset(0, 9).Value = UserForm1.TextBox7.Value
    
    'página Venda
    If UserForm1.OptionButton11.Value = True Then
        range1.Offset(0, 10).Value = "Comprou"
    ElseIf UserForm1.OptionButton12.Value = True Then
        range1.Offset(0, 10).Value = "Não comprou"
    End If
    
End Sub


Private Sub UserForm_Initialize()

    UserForm1.TextBox8.Visible = False
    UserForm1.Frame1.Visible = False
    UserForm1.Frame2.Visible = False
    UserForm1.Frame3.Visible = False
    
    
    Label8.Caption = "Quantidade de caracteres: "
    
    
    UserForm1.Frame1.Caption = "Data"
    
    
    UserForm1.OptionButton1.Caption = "Facebook"
    UserForm1.OptionButton4.Caption = "Zap"
    UserForm1.OptionButton5.Caption = "OLX"
    UserForm1.OptionButton6.Caption = "Outros"
   
    
    
    UserForm1.OptionButton2.Caption = "Sim"
    UserForm1.OptionButton3.Caption = "Não"
    
    
    UserForm1.OptionButton7.Caption = "Lançamento"
    UserForm1.OptionButton8.Caption = "Usado"
    
    UserForm1.OptionButton9.Caption = "Potencial"
    UserForm1.OptionButton10.Caption = "Pesquisando"
    UserForm1.OptionButton13.Caption = "Frio"
    
    UserForm1.OptionButton11.Caption = "Comprou"
    UserForm1.OptionButton12.Caption = "Não comprou"
    
    
    
   
End Sub



