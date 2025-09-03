' ============================================================
' CÓDIGO DO USERFORM frmSenha - Para entrada mascarada de senha
' ============================================================

' INSTRUÇÕES PARA CRIAR O USERFORM:
'
' 1. No VBA Editor, clique com botão direito no projeto
' 2. Selecione "Insert" > "UserForm"
' 3. Configure as propriedades do UserForm:
'    - (Name): frmSenha
'    - Caption: Entrada de Senha
'    - Height: 160
'    - Width: 320
'    - ShowModal: True
'    - StartUpPosition: 1 - CenterOwner
'
' 4. Adicione os controles:
'    a) Label (lblPrompt):
'       - Caption: Digite a senha:
'       - Left: 12
'       - Top: 12
'       - Width: 280
'       - Height: 15
'
'    b) TextBox (txtSenha):
'       - (Name): txtSenha
'       - Left: 12
'       - Top: 36
'       - Width: 280
'       - Height: 21
'       - PasswordChar: *
'       - TabIndex: 0
'
'    c) CommandButton (cmdOK):
'       - (Name): cmdOK
'       - Caption: &OK
'       - Left: 130
'       - Top: 75
'       - Width: 75
'       - Height: 25
'       - Default: True
'       - TabIndex: 1
'
'    d) CommandButton (cmdCancelar):
'       - (Name): cmdCancelar
'       - Caption: &Cancelar
'       - Left: 215
'       - Top: 75
'       - Width: 75
'       - Height: 25
'       - Cancel: True
'       - TabIndex: 2
'
' 5. Cole o código abaixo no UserForm:

Option Explicit

' Variáveis públicas para comunicação
Public SenhaInformada As String
Public Cancelado As Boolean

' Evento do botão OK
Private Sub cmdOK_Click()
    
    ' Validar se foi digitada alguma senha
    If Trim(txtSenha.Text) = "" Then
        MsgBox "Por favor, digite uma senha.", vbExclamation, "Senha necessária"
        txtSenha.SetFocus
        Exit Sub
    End If
    
    ' Definir valores de retorno
    SenhaInformada = txtSenha.Text
    Cancelado = False
    
    ' Fechar o formulário
    Me.Hide
    
End Sub

' Evento do botão Cancelar
Private Sub cmdCancelar_Click()
    
    ' Definir valores de retorno para cancelamento
    SenhaInformada = ""
    Cancelado = True
    
    ' Fechar o formulário
    Me.Hide
    
End Sub

' Evento de inicialização do formulário
Private Sub UserForm_Initialize()
    
    ' Configurações iniciais
    Cancelado = True
    SenhaInformada = ""
    
    ' Focar no campo de senha
    txtSenha.SetFocus
    
End Sub

' Evento quando usuário tenta fechar com X
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Se fechou pelo X (menu de controle)
    If CloseMode = vbFormControlMenu Then
        Cancelado = True
        SenhaInformada = ""
    End If
    
End Sub

' Evento de ativação do formulário
Private Sub UserForm_Activate()
    
    ' Garantir foco no campo de senha
    txtSenha.SetFocus
    
End Sub

' Evento KeyPress do TextBox da senha (opcional - para melhorar UX)
Private Sub txtSenha_KeyPress(ByVal KeyAscii As Integer)
    
    ' Se pressionou Enter, acionar botão OK
    If KeyAscii = 13 Then  ' Enter
        KeyAscii = 0  ' Cancelar o beep
        cmdOK_Click
    End If
    
    ' Se pressionou Escape, acionar botão Cancelar
    If KeyAscii = 27 Then  ' Escape
        KeyAscii = 0
        cmdCancelar_Click
    End If
    
End Sub

' ============================================================
' LAYOUT VISUAL ALTERNATIVO MAIS MODERNO (OPCIONAL)
' ============================================================

' Se quiser um visual mais moderno, configure assim:
'
' UserForm Properties:
' - BackColor: &H00F0F0F0& (cinza claro)
' - BorderStyle: 1 - fmBorderStyleSingle
' - Caption: 🔒 Entrada de Senha Segura
'
' Label (lblPrompt):
' - Font: Segoe UI, 9pt
' - ForeColor: &H00404040& (cinza escuro)
'
' TextBox (txtSenha):
' - Font: Segoe UI, 10pt
' - BackColor: &H00FFFFFF& (branco)
' - BorderStyle: 1 - fmBorderStyleSingle
'
' Buttons:
' - Font: Segoe UI, 9pt
' - BackColor: &H00E0E0E0& (cinza)
' - cmdOK BackColor: &H0000AA00& (verde) para destacar
' - ForeColor do OK: &H00FFFFFF& (branco)

' ============================================================
' VERSÃO COMPACTA PARA COPIAR DIRETO NO USERFORM
' ============================================================

' Option Explicit
' Public SenhaInformada As String
' Public Cancelado As Boolean
'
' Private Sub cmdOK_Click()
'     If Trim(txtSenha.Text) = "" Then
'         MsgBox "Digite uma senha.", vbExclamation: txtSenha.SetFocus: Exit Sub
'     End If
'     SenhaInformada = txtSenha.Text: Cancelado = False: Me.Hide
' End Sub
'
' Private Sub cmdCancelar_Click()
'     SenhaInformada = "": Cancelado = True: Me.Hide
' End Sub
'
' Private Sub UserForm_Initialize()
'     Cancelado = True: SenhaInformada = "": txtSenha.SetFocus
' End Sub
'
' Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'     If CloseMode = vbFormControlMenu Then Cancelado = True: SenhaInformada = ""
' End Sub
'
' Private Sub txtSenha_KeyPress(ByVal KeyAscii As Integer)
'     If KeyAscii = 13 Then KeyAscii = 0: cmdOK_Click
'     If KeyAscii = 27 Then KeyAscii = 0: cmdCancelar_Click
' End Sub