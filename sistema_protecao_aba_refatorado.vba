' ============================================================
' SISTEMA DE PROTEÇÃO DE ABA COM SENHA MASCARADA - REFATORADO
' Versão Corrigida - Resolve problemas de persistência da senha
' ============================================================

Option Explicit

' ==================================================
' CONFIGURAÇÕES PRINCIPAIS - AJUSTE AQUI
' ==================================================
Const NOME_ABA As String = "Consolidado NF+SE"  ' Nome da aba a ser protegida
Const NOME_BOTAO As String = "btnBloquearDesbloquear"  ' Nome do botão
Const INICIAR_BLOQUEADO As Boolean = True  ' True = sempre inicia bloqueado quando já tem senha
Const CHAVE_SENHA As String = "SistemaProtecaoSenha"  ' Chave para salvar senha

' Variáveis globais - Mantidas durante toda a sessão
Public SenhaDefinida As String
Public AbaProtegida As Boolean
Public SenhaPersistente As String  ' Nova variável para manter senha carregada

' ==================================================
' FUNÇÃO InputSenha - SUBSTITUI InputBox com máscara
' ==================================================
Function InputSenha(Optional titulo As String = "Senha necessária", _
                    Optional mensagem As String = "Digite a senha:") As String
    
    ' Verificar se o UserForm existe
    On Error GoTo ErroUserForm
    
    With frmSenha
        .Caption = titulo
        .lblPrompt.Caption = mensagem
        .txtSenha.Text = ""
        .Cancelado = True
        .Show
        
        If .Cancelado Then
            InputSenha = ""
        Else
            InputSenha = .SenhaInformada
        End If
    End With
    
    Unload frmSenha
    Exit Function
    
ErroUserForm:
    ' Fallback para InputBox se UserForm não existir
    MsgBox "ATENÇÃO: UserForm de senha não encontrado." & vbCrLf & _
           "Será usado método alternativo (senha visível)." & vbCrLf & vbCrLf & _
           "Para habilitar entrada mascarada, crie o UserForm 'frmSenha'.", _
           vbExclamation, "UserForm não encontrado"
    
    InputSenha = InputBox(mensagem & vbCrLf & vbCrLf & _
                         "ATENÇÃO: Senha será visível durante digitação!", titulo, "")
End Function

' ==================================================
' FUNÇÃO PRINCIPAL - ALTERNAR PROTEÇÃO DA ABA
' ==================================================
Sub AlternarProtecaoAba()
    
    Dim ws As Worksheet
    Dim btn As Button
    
    ' Verificar se a aba existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "ERRO: A aba '" & NOME_ABA & "' deve existir na planilha." & vbCrLf & vbCrLf & _
               "• Verifique se o nome da aba não foi alterado" & vbCrLf & _
               "• Esta aba é necessária para o funcionamento do sistema", vbCritical, "Aba Obrigatória"
        Exit Sub
    End If
    
    ' Garantir que estrutura da pasta nunca seja protegida
    If ThisWorkbook.ProtectStructure Then
        On Error Resume Next
        ThisWorkbook.Unprotect
        On Error GoTo 0
    End If
    
    ' Inicializar senha se necessário
    InicializarSenha
    
    ' Tentar encontrar o botão existente
    On Error Resume Next
    Set btn = ws.Buttons(NOME_BOTAO)
    On Error GoTo 0
    
    ' Se não existe botão, criar um novo
    If btn Is Nothing Then
        Set btn = CriarBotaoSeguro(ws)
        If btn Is Nothing Then
            CriarBotaoComTexto ws, IIf(ws.ProtectContents, "Desbloquear", "Bloquear")
            On Error Resume Next
            Set btn = ws.Buttons(NOME_BOTAO)
            On Error GoTo 0
        End If
    End If
    
    ' Verificar se botão é válido e recriar se necessário
    On Error Resume Next
    Dim testeBotao As String
    testeBotao = btn.Caption
    On Error GoTo 0
    
    If testeBotao = "" Then
        CriarBotaoComTexto ws, IIf(ws.ProtectContents, "Desbloquear", "Bloquear")
        On Error Resume Next
        Set btn = ws.Buttons(NOME_BOTAO)
        On Error GoTo 0
    End If
    
    ' Verificar estado atual da proteção e alternar
    If ws.ProtectContents Then
        ' Aba está protegida - DESBLOQUEAR
        DesbloquearAba ws, btn
    Else
        ' Aba está desprotegida - BLOQUEAR
        BloquearAba ws, btn
    End If
    
End Sub

' ==================================================
' INICIALIZAR SENHA - Nova função para resolver persistência
' ==================================================
Sub InicializarSenha()
    
    ' Se ainda não carregou a senha, tentar carregar
    If SenhaDefinida = "" And SenhaPersistente = "" Then
        SenhaPersistente = CarregarSenha()
        SenhaDefinida = SenhaPersistente
    End If
    
    ' Se não encontrou senha salva, definir como vazio para primeira execução
    If SenhaPersistente = "" Then
        SenhaDefinida = ""
    Else
        SenhaDefinida = SenhaPersistente
    End If
    
End Sub

' ==================================================
' BLOQUEAR ABA - Corrigido para persistir senha
' ==================================================
Sub BloquearAba(ws As Worksheet, btn As Button)
    
    Dim senhaParaUsar As String
    
    ' Garantir que a senha esteja inicializada
    InicializarSenha
    
    ' Verificar se é a primeira execução (senha ainda não foi definida)
    If SenhaDefinida = "" Or SenhaPersistente = "" Then
        senhaParaUsar = SolicitarCriacaoSenha()
        If senhaParaUsar = "" Then
            Exit Sub  ' Usuário cancelou
        End If
        ' Salvar nos dois lugares para garantir persistência
        SenhaDefinida = senhaParaUsar
        SenhaPersistente = senhaParaUsar
        SalvarSenha senhaParaUsar
    Else
        senhaParaUsar = SenhaDefinida
    End If
    
    On Error GoTo ErroBloquear
    
    ' REMOVER o botão antes de proteger
    On Error Resume Next
    ws.Buttons(NOME_BOTAO).Delete
    On Error GoTo 0
    
    ' Proteger a planilha
    ws.Protect Password:=senhaParaUsar, _
              DrawingObjects:=False, _
              Contents:=True, _
              Scenarios:=True, _
              UserInterfaceOnly:=False, _
              AllowFormattingCells:=False, _
              AllowFormattingColumns:=False, _
              AllowFormattingRows:=False, _
              AllowInsertingColumns:=False, _
              AllowInsertingRows:=False, _
              AllowInsertingHyperlinks:=False, _
              AllowDeletingColumns:=False, _
              AllowDeletingRows:=False, _
              AllowSorting:=True, _
              AllowFiltering:=True, _
              AllowUsingPivotTables:=True
    
    ' RECRIAR o botão com texto correto
    CriarBotaoComTexto ws, "Desbloquear"
    
    ' Atualizar variável de controle
    AbaProtegida = True
    
    MsgBox "Aba '" & NOME_ABA & "' foi BLOQUEADA com sucesso!" & vbCrLf & vbCrLf & _
           "Proteções aplicadas:" & vbCrLf & _
           "• Dados protegidos contra alteração" & vbCrLf & _
           "• Exclusão e inserção de dados bloqueada" & vbCrLf & _
           "• Filtros e tabelas dinâmicas liberados" & vbCrLf & _
           "• Outras abas permanecem livres", _
           vbInformation, "Proteção Ativada"
    
    Exit Sub
    
ErroBloquear:
    MsgBox "Não foi possível bloquear a aba neste momento." & vbCrLf & _
           "Tente novamente ou contate o suporte técnico.", vbCritical, "Erro no Bloqueio"
    
End Sub

' ==================================================
' DESBLOQUEAR ABA - Corrigido para usar senha persistente
' ==================================================
Sub DesbloquearAba(ws As Worksheet, btn As Button)
    
    Dim senhaInformada As String
    Dim tentativas As Integer
    
    ' Garantir que a senha esteja carregada
    InicializarSenha
    
    ' Se ainda não tem senha definida, algo está errado
    If SenhaDefinida = "" And SenhaPersistente = "" Then
        MsgBox "ERRO: Não foi possível encontrar a senha de proteção." & vbCrLf & _
               "A aba pode ter sido protegida externamente." & vbCrLf & vbCrLf & _
               "Use 'Redefinir Senha' para recriar o sistema de proteção.", _
               vbCritical, "Senha não encontrada"
        Exit Sub
    End If
    
    ' Loop com limite de tentativas para maior segurança
    tentativas = 0
    Do While tentativas < 3
        ' Solicitar senha do usuário
        senhaInformada = InputSenha("Desbloqueio necessário", _
                                  "Digite a senha para desbloquear a aba '" & NOME_ABA & "':")
        
        ' Verificar se o usuário cancelou
        If senhaInformada = "" Then
            MsgBox "Operação de desbloqueio cancelada pelo usuário.", vbInformation, "Operação Cancelada"
            Exit Sub
        End If
        
        ' Verificar senha
        If senhaInformada = SenhaDefinida Then
            Exit Do  ' Senha correta, sair do loop
        Else
            tentativas = tentativas + 1
            If tentativas >= 3 Then
                MsgBox "Muitas tentativas de senha incorreta." & vbCrLf & _
                       "Por segurança, a operação será cancelada." & vbCrLf & vbCrLf & _
                       "Tente novamente mais tarde.", vbCritical, "Limite de tentativas excedido"
                Exit Sub
            End If
            
            MsgBox "Senha incorreta! (" & tentativas & "/3 tentativas)" & vbCrLf & _
                   "Verifique a senha e tente novamente.", vbExclamation, "Senha Incorreta"
        End If
    Loop
    
    On Error GoTo ErroDesbloquear
    
    ' REMOVER o botão antes de desproteger
    On Error Resume Next
    ws.Buttons(NOME_BOTAO).Delete
    On Error GoTo 0
    
    ' Desproteger a planilha
    ws.Unprotect Password:=SenhaDefinida
    
    ' RECRIAR o botão com texto correto
    CriarBotaoComTexto ws, "Bloquear"
    
    ' Atualizar variável de controle
    AbaProtegida = False
    
    MsgBox "Aba '" & NOME_ABA & "' foi DESBLOQUEADA com sucesso!" & vbCrLf & vbCrLf & _
           "Acesso liberado:" & vbCrLf & _
           "• Edição de dados permitida" & vbCrLf & _
           "• Inserção e exclusão liberadas" & vbCrLf & _
           "• Formatação desbloqueada" & vbCrLf & _
           "• Acesso total restaurado", _
           vbInformation, "Proteção Removida"
    
    Exit Sub
    
ErroDesbloquear:
    MsgBox "Não foi possível desbloquear a aba neste momento." & vbCrLf & _
           "Verifique se a senha está correta e tente novamente." & vbCrLf & _
           "Se o problema persistir, contate o suporte técnico.", vbCritical, "Erro no Desbloqueio"
    
End Sub

' ==================================================
' CRIAR E GERENCIAR BOTÕES
' ==================================================

Sub CriarBotaoComTexto(ws As Worksheet, texto As String)
    
    Dim btn As Button
    
    On Error Resume Next
    
    ' Criar o botão
    Set btn = ws.Buttons.Add(10, 2, 80, ws.Rows(1).Height - 2)
    
    If Not btn Is Nothing Then
        btn.Name = NOME_BOTAO
        btn.Caption = texto
        btn.OnAction = "AlternarProtecaoAba"
        btn.Font.Size = 9
        btn.Font.Bold = True
    End If
    
    On Error GoTo 0
    
End Sub

Function CriarBotaoSeguro(ws As Worksheet) As Button
    
    Dim btn As Button
    Dim textoInicial As String
    
    Set CriarBotaoSeguro = Nothing
    
    On Error GoTo ErroCriarBotao
    
    ' Remover botão existente se houver
    On Error Resume Next
    ws.Buttons(NOME_BOTAO).Delete
    On Error GoTo 0
    
    ' Definir texto inicial baseado no estado atual
    If ws.ProtectContents Then
        textoInicial = "Desbloquear"
    Else
        textoInicial = "Bloquear"
    End If
    
    ' Aguardar processamento
    Application.Wait (Now + TimeValue("0:00:01"))
    
    ' Criar o botão
    Set btn = ws.Buttons.Add(10, 2, 80, ws.Rows(1).Height - 2)
    
    ' Configurar propriedades
    btn.Name = NOME_BOTAO
    btn.Caption = textoInicial
    btn.OnAction = "AlternarProtecaoAba"
    
    ' Configurar fonte
    On Error Resume Next
    btn.Font.Size = 9
    btn.Font.Bold = True
    On Error GoTo 0
    
    Set CriarBotaoSeguro = btn
    Exit Function
    
ErroCriarBotao:
    Set CriarBotaoSeguro = Nothing
    
End Function

' ==================================================
' GERENCIAMENTO DE SENHAS - Corrigido para múltiplos métodos
' ==================================================

Function SolicitarCriacaoSenha() As String
    
    Dim senha1 As String, senha2 As String
    
    ' Explicar ao usuário
    If MsgBox("CRIAÇÃO DE NOVA SENHA" & vbCrLf & vbCrLf & _
              "Para proteger esta aba, você precisa criar uma senha de segurança." & vbCrLf & _
              "Esta senha será salva no arquivo e usada para bloquear/desbloquear a aba." & vbCrLf & vbCrLf & _
              "IMPORTANTE: Anote esta senha em local seguro!" & vbCrLf & _
              "Sem ela, não será possível desbloquear a aba posteriormente." & vbCrLf & vbCrLf & _
              "Deseja continuar com a criação da senha?", vbQuestion + vbYesNo, "Criação de Senha de Segurança") = vbNo Then
        SolicitarCriacaoSenha = ""
        Exit Function
    End If
    
    ' Loop para criação da senha com confirmação
    Do
        ' Solicitar senha
        senha1 = InputSenha("Nova Senha de Segurança", _
                           "Crie uma senha segura para proteger a aba:" & vbCrLf & vbCrLf & _
                           "Dicas para uma senha forte:" & vbCrLf & _
                           "• Use pelo menos 4 caracteres" & vbCrLf & _
                           "• Combine letras, números e símbolos" & vbCrLf & _
                           "• Evite palavras óbvias ou dados pessoais")
        
        ' Verificar cancelamento
        If senha1 = "" Then
            If MsgBox("Deseja cancelar a criação da senha?" & vbCrLf & _
                     "Cancelar impedirá a proteção da aba.", vbQuestion + vbYesNo, "Confirmar Cancelamento") = vbYes Then
                SolicitarCriacaoSenha = ""
                Exit Function
            Else
                GoTo ProximaTentativa
            End If
        End If
        
        ' Validar tamanho mínimo
        If Len(senha1) < 4 Then
            MsgBox "Senha muito curta!" & vbCrLf & _
                   "Use pelo menos 4 caracteres para maior segurança.", vbExclamation, "Senha Inválida"
            GoTo ProximaTentativa
        End If
        
        ' Solicitar confirmação
        senha2 = InputSenha("Confirmação de Senha", _
                           "Confirme a senha digitada:" & vbCrLf & _
                           "Digite novamente a mesma senha para confirmação.")
        
        ' Verificar confirmação
        If senha2 = "" Then
            If MsgBox("Deseja cancelar a criação da senha?", vbQuestion + vbYesNo, "Cancelar") = vbYes Then
                SolicitarCriacaoSenha = ""
                Exit Function
            Else
                GoTo ProximaTentativa
            End If
        End If
        
        ' Verificar se coincidem
        If senha1 = senha2 Then
            MsgBox "Senha criada com sucesso!" & vbCrLf & _
                   "A senha foi salva no arquivo e a aba será bloqueada automaticamente.", _
                   vbInformation, "Senha Criada"
            SolicitarCriacaoSenha = senha1
            Exit Function
        Else
            MsgBox "As senhas digitadas não coincidem!" & vbCrLf & _
                   "Tente novamente.", vbExclamation, "Senhas Diferentes"
        End If
        
ProximaTentativa:
    Loop
    
End Function

' ==================================================
' SALVAR/CARREGAR SENHA - Múltiplos métodos para garantir persistência
' ==================================================

Sub SalvarSenha(senha As String)
    
    Dim sucesso As Boolean
    sucesso = False
    
    ' Método 1: CustomDocumentProperties (principal)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(CHAVE_SENHA).Delete
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties.Add CHAVE_SENHA, False, msoPropertyTypeString, senha
    If Err.Number = 0 Then sucesso = True
    On Error GoTo 0
    
    ' Método 2: Names (alternativo - mais confiável)
    On Error Resume Next
    ThisWorkbook.Names(CHAVE_SENHA).Delete
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=CHAVE_SENHA, RefersTo:=Chr(34) & senha & Chr(34)
    On Error GoTo 0
    
    ' Método 3: Planilha oculta (backup)
    CriarPlanilhaSenhaSeNecessario
    On Error Resume Next
    Dim wsBackup As Worksheet
    Set wsBackup = ThisWorkbook.Worksheets("_SistemaConfig")
    If Not wsBackup Is Nothing Then
        wsBackup.Cells(1, 1).Value = senha
    End If
    On Error GoTo 0
    
    ' Salvar arquivo para garantir persistência
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
End Sub

Function CarregarSenha() As String
    
    Dim senha As String
    senha = ""
    
    ' Método 1: Tentar CustomDocumentProperties
    On Error Resume Next
    senha = ThisWorkbook.CustomDocumentProperties(CHAVE_SENHA).Value
    On Error GoTo 0
    
    ' Método 2: Se não encontrou, tentar Names
    If senha = "" Then
        On Error Resume Next
        Dim nomeObj As Name
        Set nomeObj = ThisWorkbook.Names(CHAVE_SENHA)
        If Not nomeObj Is Nothing Then
            senha = Replace(Replace(nomeObj.RefersTo, "=""", ""), """", "")
        End If
        On Error GoTo 0
    End If
    
    ' Método 3: Se ainda não encontrou, tentar planilha backup
    If senha = "" Then
        On Error Resume Next
        Dim wsBackup As Worksheet
        Set wsBackup = ThisWorkbook.Worksheets("_SistemaConfig")
        If Not wsBackup Is Nothing Then
            senha = wsBackup.Cells(1, 1).Value
        End If
        On Error GoTo 0
    End If
    
    CarregarSenha = senha
    
End Function

' Criar planilha oculta para backup da senha (método mais confiável)
Sub CriarPlanilhaSenhaSeNecessario()
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_SistemaConfig")
    On Error GoTo 0
    
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add
        If Not ws Is Nothing Then
            ws.Name = "_SistemaConfig"
            ws.Visible = xlSheetVeryHidden
            ws.Cells(1, 1).Value = ""  ' Placeholder para senha
        End If
        On Error GoTo 0
    End If
    
End Sub

' ==================================================
' REDEFINIR SENHA - Corrigido
' ==================================================
Sub RedefinirSenha()
    
    Dim novaSenha As String
    Dim senhaAtual As String
    Dim ws As Worksheet
    
    ' Verificar se a aba existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "ERRO: A aba '" & NOME_ABA & "' não foi encontrada.", vbCritical, "Aba não encontrada"
        Exit Sub
    End If
    
    ' Carregar senha atual
    senhaAtual = CarregarSenha()
    
    If senhaAtual <> "" Then
        ' Solicitar senha atual primeiro
        Dim senhaVerificacao As String
        senhaVerificacao = InputSenha("Verificação de Senha Atual", _
                                    "Para redefinir a senha, digite primeiro a senha atual:")
        
        If senhaVerificacao = "" Then
            MsgBox "Operação cancelada.", vbInformation, "Cancelado"
            Exit Sub
        End If
        
        If senhaVerificacao <> senhaAtual Then
            MsgBox "Senha atual incorreta!" & vbCrLf & _
                   "A operação foi cancelada por segurança.", vbCritical, "Acesso Negado"
            Exit Sub
        End If
    End If
    
    ' Solicitar nova senha
    novaSenha = SolicitarCriacaoSenha()
    
    If novaSenha <> "" Then
        ' Atualizar todas as variáveis e métodos de armazenamento
        SenhaDefinida = novaSenha
        SenhaPersistente = novaSenha
        SalvarSenha novaSenha
        
        ' Se a aba estava protegida, re-proteger com nova senha
        If ws.ProtectContents Then
            On Error Resume Next
            ws.Unprotect Password:=senhaAtual
            ws.Protect Password:=novaSenha, UserInterfaceOnly:=False
            On Error GoTo 0
        End If
        
        MsgBox "Senha redefinida com sucesso!" & vbCrLf & _
               "A nova senha já está ativa e foi salva no arquivo.", vbInformation, "Senha Atualizada"
    End If
    
End Sub

' ==================================================
' VERIFICAÇÃO E INICIALIZAÇÃO - Corrigido
' ==================================================

Sub VerificarEstadoInicial()
    
    Dim ws As Worksheet
    Dim btn As Button
    
    ' Garantir que estrutura da pasta nunca seja protegida
    If ThisWorkbook.ProtectStructure Then
        On Error Resume Next
        ThisWorkbook.Unprotect
        On Error GoTo 0
    End If
    
    ' Verificar se a aba existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Inicializar sistema de senhas
    InicializarSenha
    
    ' Verificar se deve iniciar bloqueado
    If INICIAR_BLOQUEADO And Not ws.ProtectContents And SenhaDefinida <> "" Then
        On Error Resume Next
        ws.Protect Password:=SenhaDefinida, _
                  DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                  UserInterfaceOnly:=False, AllowFormattingCells:=False, _
                  AllowFormattingColumns:=False, AllowFormattingRows:=False, _
                  AllowInsertingColumns:=False, AllowInsertingRows:=False, _
                  AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _
                  AllowDeletingRows:=False, AllowSorting:=True, _
                  AllowFiltering:=True, AllowUsingPivotTables:=True
        On Error GoTo 0
    End If
    
    ' Atualizar variável de controle
    AbaProtegida = ws.ProtectContents
    
    ' Verificar/criar botão
    On Error Resume Next
    Set btn = ws.Buttons(NOME_BOTAO)
    On Error GoTo 0
    
    If btn Is Nothing Then
        Set btn = CriarBotaoSeguro(ws)
    End If
    
    ' Atualizar texto do botão
    If Not btn Is Nothing Then
        On Error Resume Next
        If AbaProtegida Then
            btn.Caption = "Desbloquear"
        Else
            btn.Caption = "Bloquear"
        End If
        On Error GoTo 0
    End If
    
End Sub

' ==================================================
' EVENTOS AUTOMÁTICOS
' ==================================================

' Este código deve ir no módulo ThisWorkbook
Private Sub Workbook_Open()
    Application.OnTime Now + TimeValue("00:00:02"), "VerificarEstadoInicial"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Garantir que a senha seja salva antes de fechar
    If SenhaDefinida <> "" Then
        SalvarSenha SenhaDefinida
    End If
End Sub

' ==================================================
' UTILITÁRIOS E DIAGNÓSTICO
' ==================================================

Function AbaEstaProtegida() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        AbaEstaProtegida = ws.ProtectContents
    Else
        AbaEstaProtegida = False
    End If
End Function

Sub RemoverBotao()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    If Not ws Is Nothing Then
        ws.Buttons(NOME_BOTAO).Delete
    End If
    On Error GoTo 0
End Sub

' Função de diagnóstico para verificar estado do sistema
Sub DiagnosticarSistema()
    
    Dim msg As String
    
    msg = "DIAGNÓSTICO DO SISTEMA DE PROTEÇÃO" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    
    ' Verificar aba
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    On Error GoTo 0
    
    If ws Is Nothing Then
        msg = msg & "❌ Aba '" & NOME_ABA & "' não encontrada!" & vbCrLf
    Else
        msg = msg & "✅ Aba '" & NOME_ABA & "' encontrada" & vbCrLf
        msg = msg & "   └ Status: " & IIf(ws.ProtectContents, "PROTEGIDA", "DESPROTEGIDA") & vbCrLf
    End If
    
    ' Verificar senhas
    InicializarSenha
    msg = msg & vbCrLf & "SENHAS:" & vbCrLf
    msg = msg & "   • SenhaDefinida: " & IIf(SenhaDefinida <> "", "✅ Definida", "❌ Vazia") & vbCrLf
    msg = msg & "   • SenhaPersistente: " & IIf(SenhaPersistente <> "", "✅ Carregada", "❌ Vazia") & vbCrLf
    msg = msg & "   • Senha Salva: " & IIf(CarregarSenha() <> "", "✅ Encontrada", "❌ Não encontrada") & vbCrLf
    
    ' Verificar botão
    On Error Resume Next
    Dim btn As Button
    Set btn = ws.Buttons(NOME_BOTAO)
    On Error GoTo 0
    
    If btn Is Nothing Then
        msg = msg & vbCrLf & "❌ Botão não encontrado" & vbCrLf
    Else
        msg = msg & vbCrLf & "✅ Botão encontrado" & vbCrLf
        On Error Resume Next
        msg = msg & "   └ Texto: '" & btn.Caption & "'" & vbCrLf
        On Error GoTo 0
    End If
    
    MsgBox msg, vbInformation, "Diagnóstico do Sistema"
    
End Sub

' Sub para limpar completamente o sistema (uso em desenvolvimento)
Sub LimparSistemaCompleto()
    
    If MsgBox("ATENÇÃO: Esta operação irá:" & vbCrLf & vbCrLf & _
              "• Remover TODAS as senhas salvas" & vbCrLf & _
              "• Desproteger a aba completamente" & vbCrLf & _
              "• Excluir o botão de controle" & vbCrLf & _
              "• Resetar todas as variáveis" & vbCrLf & vbCrLf & _
              "Confirma a operação?", vbCritical + vbYesNo, "Limpar Sistema Completo") = vbNo Then
        Exit Sub
    End If
    
    ' Limpar variáveis
    SenhaDefinida = ""
    SenhaPersistente = ""
    AbaProtegida = False
    
    ' Remover propriedades customizadas
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(CHAVE_SENHA).Delete
    ThisWorkbook.Names(CHAVE_SENHA).Delete
    On Error GoTo 0
    
    ' Remover planilha backup
    On Error Resume Next
    Dim wsBackup As Worksheet
    Set wsBackup = ThisWorkbook.Worksheets("_SistemaConfig")
    If Not wsBackup Is Nothing Then
        Application.DisplayAlerts = False
        wsBackup.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Desproteger aba e remover botão
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA)
    If Not ws Is Nothing Then
        ws.Unprotect  ' Tentar sem senha primeiro
        ws.Buttons(NOME_BOTAO).Delete
    End If
    On Error GoTo 0
    
    MsgBox "Sistema limpo com sucesso!" & vbCrLf & _
           "Todas as senhas e configurações foram removidas.", vbInformation, "Limpeza Concluída"
    
End Sub