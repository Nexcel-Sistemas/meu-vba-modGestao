Option Explicit

Public SistemaEmColapso As Boolean

#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Const URL_ENVIO As String = "https://docs.google.com/forms/d/e/1FAIpQLScLeGEJqMGrFwuOgUfJHuvwnr4TW6CXHG6tSlO3xNMxovw3Xg/formResponse"
Const ID_NOME As String = "entry.1607168817"
Const ID_PC As String = "entry.1526859411"
Const ID_USER As String = "entry.1174285135"
Const ID_IP As String = "entry.1960945120"
Const ID_DATA As String = "entry.1222486771"
Const ID_EMAIL As String = "entry.1862125144"
Const URL_CHECK_ONLINE As String = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS_MZboJrMkb9t05Fp6p4xhqBWyV44D9E9QmKjSYO7O4PV8g94rLtbNqhy-DyRgbNhCtAKIwUKginqK/pub?output=csv"
Private Const NOME_ABA_OFFLINE As String = "Fila_Envio_Sys"

Public Function SENHA_SISTEMA() As String
    SENHA_SISTEMA = Chr(82) & Chr(51) & Chr(98) & Chr(51) & Chr(67) & Chr(99) & Chr(52) & Chr(64) & Chr(50) & Chr(50) & Chr(64) & Chr(35) & Chr(42) & Chr(42) & Chr(42)
End Function

Public Sub CONFIGURAR_VENDA()
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    On Error GoTo 0

    Dim ObjetoVBA As Object
    Dim TemAcessoVBA As Boolean
    TemAcessoVBA = False
    
    On Error Resume Next
    Set ObjetoVBA = ThisWorkbook.VBProject
    If Err.Number = 0 Then TemAcessoVBA = True
    On Error GoTo 0
    
    If Not TemAcessoVBA Then
        MsgBox "OPERAÇÃO ABORTADA!" & vbNewLine & vbNewLine & "Para compilar o sistema ocultando as macros, o Excel precisa de permissão de acesso." & vbNewLine & vbNewLine & "Vá em: Arquivo > Opções > Central de Confiabilidade > Configurações > Configurações de Macro." & vbNewLine & "Marque a caixa: 'Confiar no acesso ao modelo de objeto do projeto do VBA'.", vbCritical, "Nexcel Sênior - Auditoria de Permissão"
        Exit Sub
    End If
    
    If ThisWorkbook.VBProject.Protection = 1 Then
        MsgBox "OPERAÇÃO ABORTADA (Prevenção de Erro 50289):" & vbNewLine & vbNewLine & "O Projeto VBA da sua Matriz está PROTEGIDO COM SENHA. É impossível carimbar a ocultação das macros em um projeto trancado." & vbNewLine & vbNewLine & "SOLUÇÃO:" & vbNewLine & "1. Aperte ALT + F11" & vbNewLine & "2. Vá em Ferramentas > Propriedades do VBAProject > Proteção" & vbNewLine & "3. Desmarque 'Bloquear projeto para exibição' e apague a senha." & vbNewLine & "4. Clique OK, salve a planilha e tente novamente." & vbNewLine & vbNewLine & "Lembre-se: O arquivo do cliente será blindado via código, portanto não precisa dessa senha nativa.", vbCritical, "Nexcel Sênior - Desbloqueio Necessário"
        Exit Sub
    End If

    Dim ModoDeusAtivo As Boolean
    ModoDeusAtivo = (VBA.Dir(VBA.Environ("APPDATA") & "\admin_key.txt", vbHidden + vbSystem + vbNormal) <> "")
    
    If Not ModoDeusAtivo Then
        Dim SenhaAdmin As String
        SenhaAdmin = InputBox("ÁREA RESTRITA AO VENDEDOR." & vbNewLine & vbNewLine & "Digite a senha de administrador para configurar uma nova venda:", "Segurança do Sistema")
        If SenhaAdmin <> SENHA_SISTEMA() Then
            Call Cortina_De_Ferro("TENTATIVA DE VIOLAÇÃO:" & vbNewLine & "Senha de administrador incorreta." & vbNewLine & "Bloqueio de segurança ativado.")
            Exit Sub
        End If
    End If

    Dim Opcao As String, ws As Worksheet, CaminhoSalvar As Variant
    Dim QtdLicencas As String, i As Long
    Dim CaminhoLista As Variant, wbLista As Workbook, wsLista As Worksheet
    Dim UltimaLinhaLista As Long, vDados As Variant, TemLista As Boolean: TemLista = False
    
    Opcao = InputBox("SELECIONE O TIPO DE LICENCIAMENTO:" & vbNewLine & vbNewLine & "--- PESSOAL (Trava PC + Usuário) ---" & vbNewLine & "1 - Pessoal Padrão (1 PC)" & vbNewLine & "2 - Pessoal Múltiplo (Lista de Nomes)" & vbNewLine & "3 - Pessoal Múltiplo (Quantidade)" & vbNewLine & vbNewLine & "--- EMPRESARIAL (Trava Só PC) ---" & vbNewLine & "4 - Empresarial Padrão (1 PC)" & vbNewLine & "5 - Empresarial Múltiplo (Lista de Nomes)" & vbNewLine & "6 - Empresarial Múltiplo (Quantidade)", "Configurador - Lucas Lima")
    If Not (Opcao Like "[1-6]") Then Exit Sub
    
    Dim MaxVagas As String: MaxVagas = "1"
    Dim ModoLista As Boolean: ModoLista = False
    
    Select Case Opcao
        Case "2", "5"
            ModoLista = True
            MsgBox "IMPORTAÇÃO DE LISTA:" & vbNewLine & vbNewLine & "Selecione o arquivo Excel (.xlsx ou .xls) contendo a lista." & vbNewLine & "Coluna A: Nome da Máquina (Obrigatório)" & vbNewLine & IIf(Opcao = "2", "Coluna B: Nome de Usuário (Obrigatório para Opção 2)", ""), vbInformation, "Importar Dados"
            CaminhoLista = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "SELECIONE A PLANILHA COM A LISTA")
            If CaminhoLista = False Then Exit Sub
            Application.ScreenUpdating = False
            Set wbLista = Workbooks.Open(CaminhoLista, ReadOnly:=True)
            Set wsLista = wbLista.Worksheets(1)
            UltimaLinhaLista = wsLista.Cells(wsLista.Rows.Count, "A").End(xlUp).Row
            If UltimaLinhaLista < 2 Then
                wbLista.Close False
                MsgBox "O arquivo selecionado parece estar vazio na Coluna A.", vbCritical
                Exit Sub
            End If
            vDados = wsLista.Range("A2:B" & UltimaLinhaLista).Value
            wbLista.Close False
            TemLista = True
            Application.ScreenUpdating = True
        Case "3", "6"
            QtdLicencas = InputBox("Quantas licenças (máquinas) serão permitidas no total?", "Definir Limite")
            If Not IsNumeric(QtdLicencas) Or Val(QtdLicencas) < 1 Then Exit Sub
            MaxVagas = QtdLicencas
    End Select
    
    CaminhoSalvar = Application.GetSaveAsFilename(InitialFileName:="PLANILHA_CLIENTE_FINAL", FileFilter:="Pasta de Trabalho Habilitada para Macro (*.xlsm), *.xlsm", Title:="NEXCEL: Onde deseja salvar o sistema COMPILADO?")
    If CaminhoSalvar = False Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    
    Dim abaAntiga As Worksheet
    Set abaAntiga = ThisWorkbook.Sheets("Licenca_Sys")
    If Not abaAntiga Is Nothing Then
        abaAntiga.Name = "Del_Lic_" & Format(Now, "hhmmss")
        abaAntiga.Visible = xlSheetVisible
        abaAntiga.Delete
    End If
    Set abaAntiga = Nothing
    
    Set abaAntiga = ThisWorkbook.Sheets(NOME_ABA_OFFLINE)
    If Not abaAntiga Is Nothing Then
        abaAntiga.Name = "Del_Off_" & Format(Now, "hhmmss")
        abaAntiga.Visible = xlSheetVisible
        abaAntiga.Delete
    End If
    Set abaAntiga = Nothing
    
    If Dir(modSentinela.CaminhoSentinela, vbHidden + vbSystem) <> "" Then
        SetAttr modSentinela.CaminhoSentinela, vbNormal
        Kill modSentinela.CaminhoSentinela
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Licenca_Sys"
    ws.Visible = xlSheetVeryHidden
    ws.Range("A100").Value = Opcao
    ws.Range("A101").Value = MaxVagas
    ws.Range("A102").Value = "0"
    ws.Range("A1").Value = "AGUARDANDO_REGISTRO"
    
    If ModoLista And TemLista Then
        Dim LinhaSys As Long: LinhaSys = 1
        For i = 1 To UBound(vDados, 1)
            If Opcao = "2" Then
                If Trim(vDados(i, 1)) <> "" And Trim(vDados(i, 2)) <> "" Then
                    ws.Range("L" & LinhaSys).Value = UCase(Trim(vDados(i, 2)))
                    ws.Range("M" & LinhaSys).Value = UCase(Trim(vDados(i, 1)))
                    LinhaSys = LinhaSys + 1
                End If
            ElseIf Opcao = "5" Then
                If Trim(vDados(i, 1)) <> "" Then
                    ws.Range("L" & LinhaSys).Value = UCase(Trim(vDados(i, 1)))
                    LinhaSys = LinhaSys + 1
                End If
            End If
        Next i
    End If
    
    ThisWorkbook.Protect Password:=SENHA_SISTEMA(), Structure:=True, Windows:=True
    Call ForcarOcultarTudo
    
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    
    Dim CaminhoTempOrigem As String
    CaminhoTempOrigem = Environ("TEMP") & "\Nexcel_Pre_Compilado_" & Format(Now, "yymmddhhmmss") & ".xlsm"
    ThisWorkbook.SaveCopyAs CaminhoTempOrigem
    
    Call INJETAR_PROTECAO_VISIBILIDADE(CaminhoTempOrigem)
    Call MotorDeBlindagemIntegrado(CaminhoTempOrigem, CStr(CaminhoSalvar))
    
    Call ForcarExibirTudo
    ThisWorkbook.Saved = True
    Application.EnableEvents = True
    
    MsgBox "VENDA CONFIGURADA E BLINDADA COM SUCESSO!" & vbNewLine & vbNewLine & "A planilha 100% blindada e camuflada foi salva em:" & vbNewLine & CaminhoSalvar, vbInformation, "Nexcel Sênior - Operação Perfeita"
End Sub

Public Sub REALIZAR_MIGRACAO()
    Dim Senha As String, NovoPC As String, NovoUser As String, AntigoPC As String, NomeCli As String, EmailCli As String, IPCli As String
    Dim ws As Worksheet, TipoLic As Integer, StringEnvio As String
    
    Senha = InputBox("Digite a senha QUERO_MIGRAR para prosseguir com a migração:", "Migrar Licença")
    If Senha <> "QUERO_MIGRAR" Then MsgBox "Senha inválida.", vbCritical: Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Licenca_Sys")
    TipoLic = Val(ws.Range("A100").Value)
    AntigoPC = ws.Range("A1").Value
    NomeCli = ws.Range("H1").Value
    EmailCli = ws.Range("F1").Value
    IPCli = MeuIP()
    
    NovoPC = InputBox("Digite o NOME DO COMPUTADOR de destino:" & vbNewLine & vbNewLine & "O arquivo será bloqueado aqui e liberado lá.", "Migração - Passo 1")
    If NovoPC = "" Then Exit Sub
    
    NovoUser = ""
    If TipoLic <= 3 Then
        NovoUser = InputBox("Digite o NOME DE USUÁRIO do destino (Login do Windows):" & vbNewLine & vbNewLine & "Licença Pessoal exige Nome de Usuário.", "Migração - Passo 2")
        If NovoUser = "" Then MsgBox "Para licença pessoal, o usuário é obrigatório.", vbCritical: Exit Sub
    End If
    
    If MsgBox("CONFIRMAÇÃO DE MIGRAÇÃO" & vbNewLine & vbNewLine & "De: " & AntigoPC & vbNewLine & "Para: " & NovoPC & vbNewLine & IIf(NovoUser <> "", "Novo User: " & NovoUser & vbNewLine, "") & "Ao clicar em SIM, este computador atual será BLOQUEADO.", vbExclamation + vbYesNo, "Confirmar Migração") = vbYes Then
        Application.ScreenUpdating = False
        StringEnvio = "MIGRACAO | De: " & AntigoPC & " | Para: " & NovoPC & " [Item " & modSentinela.ID_PRODUTO & "]"
        Call EnviarDadosGoogle(NomeCli, StringEnvio, NovoUser, IPCli, Format(Now, "dd/mm/yyyy hh:mm:ss"), EmailCli)
        ws.Range("A1").Value = UCase(NovoPC)
        If NovoUser <> "" Then ws.Range("B1").Value = UCase(NovoUser)
        
        Dim TimeStamp As String
        TimeStamp = Format(Now, "yyyymmddhhmmss")
        ws.Range("C1").Value = "EM_TRANSITO|" & TimeStamp
        Call modSentinela.GravarStatusSentinela("MIGRADO_LOCAL")
        SaveSetting "SysSentinel_Security", "Bans", AntigoPC, TimeStamp
    
        Call ForcarOcultarTudo
        ThisWorkbook.Save
        Call Cortina_Migracao("AVISO DE MIGRAÇÃO: Esta licença foi transferida para o computador: " & NovoPC)
    End If
End Sub

Public Sub RegistrarAceite(ByVal NomeCli As String, ByVal EmailCliente As String, ByVal IPCli As String, ByVal MensagemTermo As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Licenca_Sys")
    Dim DataHora As String, NomePC As String, NomeUsuarioEnvio As String
    DataHora = Format(Now, "dd/mm/yyyy hh:mm:ss")
    NomePC = Environ("COMPUTERNAME")
    NomeUsuarioEnvio = Environ("USERNAME")
    ws.Range("D1").Value = "ACEITO_EM_" & DataHora
    ws.Range("E1").Value = NomeUsuarioEnvio
    ws.Range("F1").Value = EmailCliente
    ws.Range("H1").Value = NomeCli
    ws.Range("I1").Value = IPCli
    ws.Range("G1").Value = "AGUARDANDO_SENHA"
    Call AtualizarBaseVendas(NomeCli, NomePC, NomeUsuarioEnvio, IPCli, DataHora, EmailCliente, MensagemTermo)
    ThisWorkbook.Save
End Sub

Public Sub AtualizarCarimboVisual()
    Dim wsSys As Worksheet, wsInicio As Worksheet
    On Error Resume Next
    Set wsSys = ThisWorkbook.Sheets("Licenca_Sys")
    Set wsInicio = ThisWorkbook.Sheets("Início")
    If Not wsSys Is Nothing And Not wsInicio Is Nothing Then
        Dim Email As String, Data As String
        Email = wsSys.Range("F1").Value
        Data = wsSys.Range("D1").Value
        If Email <> "" Then
            wsInicio.Range("A50").Value = "LICENCIADO PARA: " & Email
            wsInicio.Range("A51").Value = "STATUS: " & Replace(Data, "_", " ")
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub VerificarPendenciasEnvio()
    Call SincronizarPendentes
End Sub

Public Sub ForcarOcultarTudo()
    Dim ws As Worksheet, wsTravado As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    Set wsTravado = ThisWorkbook.Sheets("TRAVADO")
    If wsTravado Is Nothing Then
        Set wsTravado = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsTravado.Name = "TRAVADO"
    End If
    If Not wsTravado Is Nothing Then
        With wsTravado
            .Visible = xlSheetVisible
            .Unprotect SENHA_SISTEMA()
            .Cells.Clear
            .Cells.Interior.Color = RGB(169, 208, 142)
            .Cells.Font.Color = RGB(255, 255, 255)
            Application.DisplayAlerts = False
            With .Range("B10:S14")
                .MergeCells = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Name = "Calibri"
                .Font.Size = 36
                .Font.Bold = True
                .WrapText = False
                .Value = "AGUARDANDO O CARREGAMENTO DAS MACROS..."
            End With
            Application.DisplayAlerts = True
            .Protect Password:=SENHA_SISTEMA(), UserInterfaceOnly:=True
        End With
    End If
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is Nothing Then
            If UCase(ws.Name) <> "TRAVADO" Then ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    ThisWorkbook.Protect Password:=SENHA_SISTEMA(), Structure:=True, Windows:=True
    On Error GoTo 0
    Application.ScreenUpdating = True
    DoEvents
    Sleep 150
    Set wsTravado = Nothing
    Set ws = Nothing
End Sub

Public Sub ForcarExibirTudo()
    Dim ws As Worksheet
    Dim PlanilhaLiberada As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is Nothing Then
            If UCase(ws.Name) <> "TRAVADO" And UCase(ws.Name) <> "LICENCA_SYS" And UCase(ws.Name) <> UCase(NOME_ABA_OFFLINE) Then
                ws.Visible = xlSheetVisible
                If PlanilhaLiberada Is Nothing Then Set PlanilhaLiberada = ws
            End If
        End If
    Next ws
    If Not PlanilhaLiberada Is Nothing Then PlanilhaLiberada.Activate
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets("TRAVADO")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    On Error GoTo 0
    Application.ScreenUpdating = True
    DoEvents
    Sleep 150
    Set ws = Nothing
    Set PlanilhaLiberada = Nothing
End Sub

Public Sub AplicarCortinaDeFerro(ByVal MensagemErro As String)
    Dim ws As Worksheet, wsTravado As Worksheet, shp As Shape
    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    Set wsTravado = ThisWorkbook.Sheets("TRAVADO")
    If wsTravado Is Nothing Then
        Set wsTravado = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsTravado.Name = "TRAVADO"
    End If
    If Not wsTravado Is Nothing Then
        With wsTravado
            .Visible = xlSheetVisible
            .Activate
            .Unprotect SENHA_SISTEMA()
            For Each shp In .Shapes
                shp.Delete
            Next shp
            .Cells.Clear
            .Cells.Interior.Color = RGB(150, 0, 0)
            .Cells.Font.Color = RGB(255, 255, 255)
            .Range("A1:Z100").Locked = True
            With .Range("B5:S6")
                .MergeCells = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "SISTEMA BLOQUEADO PELO FORNECEDOR."
                .Font.Size = 36
                .Font.Bold = True
            End With
            Dim LinhasMsg() As String
            Dim i As Long, linhaAtual As Long
            LinhasMsg = Split("MOTIVO DA INTERCEPTAÇÃO: " & MensagemErro, vbNewLine)
            linhaAtual = 8
            For i = LBound(LinhasMsg) To UBound(LinhasMsg)
                If Trim(LinhasMsg(i)) <> "" Then
                    With .Range(.Cells(linhaAtual, 2), .Cells(linhaAtual + 1, 19))
                        .MergeCells = True
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Value = LinhasMsg(i)
                        .Font.Size = 16
                        .Font.Bold = True
                    End With
                    linhaAtual = linhaAtual + 3
                End If
            Next i
            With .Range(.Cells(linhaAtual, 2), .Cells(linhaAtual + 1, 19))
                .MergeCells = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Value = "Esta tentativa de acesso não foi autorizada. Entre em contato com o suporte"
                .Font.Size = 14
            End With
            .Protect Password:=SENHA_SISTEMA(), UserInterfaceOnly:=True
        End With
    End If
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) <> "TRAVADO" Then ws.Visible = xlSheetVeryHidden
    Next ws
    ThisWorkbook.Protect Password:=SENHA_SISTEMA(), Structure:=True, Windows:=True
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.Visible = True
    DoEvents
    Sleep 150
End Sub

Public Sub AplicarCortinaLaranja(ByVal MensagemErro As String)
    Dim ws As Worksheet, wsTravado As Worksheet, shp As Shape
    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA_SISTEMA()
    Set wsTravado = ThisWorkbook.Sheets("TRAVADO")
    If wsTravado Is Nothing Then
        Set wsTravado = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsTravado.Name = "TRAVADO"
    End If
    If Not wsTravado Is Nothing Then
        With wsTravado
            .Visible = xlSheetVisible
            .Activate
            .Unprotect SENHA_SISTEMA()
            For Each shp In .Shapes
                shp.Delete
            Next shp
            .Cells.Clear
            .Cells.Interior.Color = RGB(226, 107, 10)
            .Cells.Font.Color = RGB(255, 255, 255)
            .Range("A1:Z100").Locked = True
            With .Range("B5:S6")
                .MergeCells = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "SISTEMA MIGRADO POR SEGURANÇA"
                .Font.Size = 36
                .Font.Bold = True
            End With
            Dim LinhasMsg() As String
            Dim i As Long, linhaAtual As Long
            LinhasMsg = Split(MensagemErro, vbNewLine)
            linhaAtual = 8
            For i = LBound(LinhasMsg) To UBound(LinhasMsg)
                If Trim(LinhasMsg(i)) <> "" Then
                    With .Range(.Cells(linhaAtual, 2), .Cells(linhaAtual + 1, 19))
                        .MergeCells = True
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Value = LinhasMsg(i)
                        .Font.Size = 16
                        .Font.Bold = True
                    End With
                    linhaAtual = linhaAtual + 3
                End If
            Next i
            .Protect Password:=SENHA_SISTEMA(), UserInterfaceOnly:=True
        End With
    End If
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) <> "TRAVADO" Then ws.Visible = xlSheetVeryHidden
    Next ws
    ThisWorkbook.Protect Password:=SENHA_SISTEMA(), Structure:=True, Windows:=True
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.Visible = True
    DoEvents
    Sleep 150
End Sub

Public Sub ExecutarBloqueioMortal()
    On Error Resume Next
    SistemaEmColapso = True
    Call AplicarCortinaDeFerro("COMANDO DE BLOQUEIO RECEBIDO.")
    Application.DisplayFullScreen = True
    AppActivate Application.Caption
    DoEvents
    MsgBox "SISTEMA BLOQUEADO PELO FORNECEDOR." & vbNewLine & "ENTRE EM CONTATO COM O SUPORTE.", vbCritical + vbSystemModal, "ACESSO NEGADO"
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    Call FecharSistemaBlindado
    On Error GoTo 0
End Sub

Public Sub Cortina_De_Ferro(ByVal Msg As String)
    On Error Resume Next
    SistemaEmColapso = True
    Call AplicarCortinaDeFerro(Msg)
    Application.DisplayFullScreen = True
    AppActivate Application.Caption
    DoEvents
    MsgBox "ACESSO NEGADO!" & vbNewLine & vbNewLine & Msg, vbCritical + vbSystemModal, "BLOQUEIO DE SEGURANÇA"
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    Call FecharSistemaBlindado
    On Error GoTo 0
End Sub

Public Sub Cortina_Migracao(ByVal Msg As String)
    On Error Resume Next
    SistemaEmColapso = True
    Call AplicarCortinaLaranja(Msg)
    Application.DisplayFullScreen = True
    AppActivate Application.Caption
    DoEvents
    MsgBox "SISTEMA MIGRADO!" & vbNewLine & vbNewLine & Msg, vbExclamation + vbSystemModal, "MIGRAÇÃO CONCLUÍDA"
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    Call FecharSistemaBlindado
    On Error GoTo 0
End Sub

Public Function VerificarStatusOnlineCSV() As String
    VerificarStatusOnlineCSV = ""
    Dim Http As Object, Stream As Object, TextoCSV As String
    Dim Linhas() As String, Colunas() As String, i As Long
    Dim LicencaCliente As String, StatusLinha As String, PC_Nuvem As String
    
    On Error Resume Next
    LicencaCliente = UCase(Trim(ThisWorkbook.Sheets("Licenca_Sys").Range("A1").Value))
    On Error GoTo 0
    
    If LicencaCliente = "" Or LicencaCliente = "AGUARDANDO_REGISTRO" Then Exit Function
    
    On Error Resume Next
    Set Http = CreateObject("MSXML2.ServerXMLHTTP")
    If Http Is Nothing Then Set Http = CreateObject("MSXML2.XMLHTTP")
    Http.SetTimeouts 5000, 5000, 5000, 5000
    
    Http.Open "GET", URL_CHECK_ONLINE & "&t=" & Int((999999 * Rnd) + 1) & Timer, False
    Http.send
    
    If Http.Status = 200 Then
        Set Stream = CreateObject("ADODB.Stream")
        Stream.Type = 1
        Stream.Open
        Stream.Write Http.responseBody
        Stream.Position = 0
        Stream.Type = 2
        Stream.Charset = "utf-8"
        TextoCSV = Stream.ReadText
        Stream.Close
        Set Stream = Nothing
        
        TextoCSV = Replace(TextoCSV, vbCr, "")
        Linhas = Split(TextoCSV, vbLf)
        
        If UBound(Linhas) > 0 Then
            For i = UBound(Linhas) To 1 Step -1
                If Trim(Linhas(i)) <> "" Then
                    Colunas = Split(Linhas(i), ",")
                    If UBound(Colunas) >= 1 Then
                        PC_Nuvem = UCase(Trim(Replace(Colunas(1), """", "")))
                        If InStr(1, PC_Nuvem, LicencaCliente) > 0 Then
                            If UBound(Colunas) >= 7 Then
                                StatusLinha = UCase(Trim(Replace(Colunas(7), """", "")))
                                If InStr(1, StatusLinha, "BLOQUEA") > 0 Or InStr(1, StatusLinha, "SUSPENS") > 0 Then
                                    VerificarStatusOnlineCSV = "BLOQUEADO"
                                    Set Http = Nothing
                                    Exit Function
                                ElseIf InStr(1, StatusLinha, "ATIVO") > 0 Or InStr(1, StatusLinha, "DESBLOQUEA") > 0 Or InStr(1, StatusLinha, "LIBERA") > 0 Then
                                    VerificarStatusOnlineCSV = "ATIVO"
                                    Set Http = Nothing
                                    Exit Function
                                End If
                            End If
                            Set Http = Nothing
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
    On Error GoTo 0
    Set Http = Nothing
End Function

Public Function VerificarContagemOnline(ByVal NomeCliente As String, ByVal EmailCliente As String) As Long
    Dim HttpReq As Object, CSVContent As String, Linhas() As String, i As Long, Contagem As Long
    If Not VerificarInternet() Then VerificarContagemOnline = 0: Exit Function
    On Error Resume Next
    Set HttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If HttpReq Is Nothing Then Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    With HttpReq
        .SetTimeouts 5000, 5000, 5000, 5000
        .Open "GET", URL_CHECK_ONLINE & "&t=" & Int((999999 * Rnd) + 1) & Timer, False
        .send
        CSVContent = .responseText
    End With
    On Error GoTo 0
    Contagem = 0
    If CSVContent <> "" Then
        CSVContent = Replace(CSVContent, vbCr, "")
        Linhas = Split(CSVContent, vbLf)
        For i = LBound(Linhas) To UBound(Linhas)
            If InStr(1, Linhas(i), NomeCliente, vbTextCompare) > 0 And InStr(1, Linhas(i), EmailCliente, vbTextCompare) > 0 And InStr(1, Linhas(i), modSentinela.ID_PRODUTO, vbTextCompare) > 0 Then
               If InStr(1, Linhas(i), "MIGRACAO", vbTextCompare) = 0 And InStr(1, Linhas(i), "MIGRADO", vbTextCompare) = 0 Then
                   Contagem = Contagem + 1
               End If
            End If
        Next i
    End If
    VerificarContagemOnline = Contagem
    Set HttpReq = Nothing
End Function

Public Sub AtualizarBaseVendas(Optional ByVal Nome As String = "", Optional ByVal PC As String = "", Optional ByVal Usuario As String = "", Optional ByVal IP As String = "", Optional ByVal DataH As String = "", Optional ByVal Email As String = "", Optional ByVal Msg As String = "")
    Dim ws As Worksheet, UltimaLinha As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Base_Vendas")
    On Error GoTo 0
    If Not ws Is Nothing Then
        UltimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        If UltimaLinha < 2 Then UltimaLinha = 2
        ws.Cells(UltimaLinha, 1).Value = DataH
        ws.Cells(UltimaLinha, 2).Value = PC
        ws.Cells(UltimaLinha, 3).Value = Usuario
        ws.Cells(UltimaLinha, 4).Value = DataH
        ws.Cells(UltimaLinha, 5).Value = Email
        ws.Cells(UltimaLinha, 6).Value = Nome
        ws.Cells(UltimaLinha, 7).Value = IP
        ws.Cells(UltimaLinha, 8).Value = Msg
    End If
    Set ws = Nothing
End Sub

Public Function URLEncodeUTF8(ByVal StringVal As String) As String
    Dim objStream As Object, Data() As Byte
    Dim i As Long, hexStr As String, res As String
    If StringVal = "" Then Exit Function
    On Error Resume Next
    Set objStream = CreateObject("ADODB.Stream")
    If objStream Is Nothing Then
        URLEncodeUTF8 = Replace(StringVal, " ", "+")
        Exit Function
    End If
    objStream.Charset = "utf-8"
    objStream.Mode = 3
    objStream.Type = 2
    objStream.Open
    objStream.WriteText StringVal
    objStream.Position = 0
    objStream.Type = 1
    Data = objStream.Read
    objStream.Close
    Set objStream = Nothing
    On Error GoTo 0
    For i = 0 To UBound(Data)
        If i < 3 And Data(i) = &HEF Then
        ElseIf i = 1 And Data(i) = &HBB And Data(0) = &HEF Then
        ElseIf i = 2 And Data(i) = &HBF And Data(0) = &HEF And Data(1) = &HBB Then
        Else
            Select Case Data(i)
                Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                    res = res & Chr(Data(i))
                Case 32
                    res = res & "+"
                Case Else
                    hexStr = Hex(Data(i))
                    If Len(hexStr) = 1 Then hexStr = "0" & hexStr
                    res = res & "%" & hexStr
            End Select
        End If
    Next i
    URLEncodeUTF8 = res
End Function

Public Function EnviarDadosGoogle(ByVal Nome As String, ByVal PC As String, ByVal Usuario As String, ByVal IP As String, ByVal Data As String, ByVal Email As String) As String
    Dim HttpReq As Object, DadosPOST As String, PC_Final As String
    If InStr(1, PC, "MIGRACAO") = 0 Then PC_Final = PC & " [Item " & modSentinela.ID_PRODUTO & "]" Else PC_Final = PC
    
    DadosPOST = ID_NOME & "=" & URLEncodeUTF8(Nome) & "&" & _
                ID_PC & "=" & URLEncodeUTF8(PC_Final) & "&" & _
                ID_USER & "=" & URLEncodeUTF8(Usuario) & "&" & _
                ID_IP & "=" & URLEncodeUTF8(IP) & "&" & _
                ID_DATA & "=" & URLEncodeUTF8(Data) & "&" & _
                ID_EMAIL & "=" & URLEncodeUTF8(Email)
                
    On Error Resume Next
    Set HttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If HttpReq Is Nothing Then Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    With HttpReq
        .SetTimeouts 5000, 5000, 5000, 5000
        .Open "POST", URL_ENVIO, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send DadosPOST
        EnviarDadosGoogle = .responseText
    End With
    On Error GoTo 0
    Set HttpReq = Nothing
End Function

Public Function MeuIP() As String
    On Error Resume Next
    If VerificarInternet() Then
        Dim HttpReq As Object
        Set HttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
        If HttpReq Is Nothing Then Set HttpReq = CreateObject("MSXML2.XMLHTTP")
        HttpReq.SetTimeouts 5000, 5000, 5000, 5000
        HttpReq.Open "GET", "https://api.ipify.org", False
        HttpReq.send
        MeuIP = HttpReq.responseText
        Set HttpReq = Nothing
    Else
        MeuIP = "Offline_Pendente"
    End If
    If Err.Number <> 0 Or MeuIP = "" Then MeuIP = "IP_Oculto"
    On Error GoTo 0
End Function

Public Function VerificarInternet() As Boolean
    Dim StatusConexao As Long
    VerificarInternet = (InternetGetConnectedState(StatusConexao, 0&) <> 0)
End Function

Public Sub SalvarPendenciaOffline(Optional ByVal Nome As String = "", Optional ByVal PC As String = "", Optional ByVal Usuario As String = "", Optional ByVal IP As String = "", Optional ByVal Data As String = "", Optional ByVal Email As String = "")
    Dim ws As Worksheet, ProxLinha As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOME_ABA_OFFLINE)
    If ws Is Nothing Then
        ThisWorkbook.Unprotect SENHA_SISTEMA()
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = NOME_ABA_OFFLINE
        ws.Range("A1:F1").Value = Array("Nome", "PC", "User", "IP", "Data", "Email")
        ws.Visible = xlSheetVeryHidden
        ThisWorkbook.Unprotect SENHA_SISTEMA()
    End If
    On Error GoTo 0
    If Not ws Is Nothing Then
        ProxLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(ProxLinha, 1).Value = Nome
        ws.Cells(ProxLinha, 2).Value = PC
        ws.Cells(ProxLinha, 3).Value = Usuario
        ws.Cells(ProxLinha, 4).Value = IP
        ws.Cells(ProxLinha, 5).Value = Data
        ws.Cells(ProxLinha, 6).Value = Email
        ThisWorkbook.Save
    End If
    Set ws = Nothing
End Sub

Public Sub SincronizarPendentes()
    Dim ws As Worksheet, i As Long, UltimaLinha As Long, Retorno As String, IP_Atualizado As String
    If Not VerificarInternet() Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOME_ABA_OFFLINE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    UltimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If UltimaLinha < 2 Then Exit Sub
    Application.ScreenUpdating = False
    IP_Atualizado = MeuIP()
    For i = UltimaLinha To 2 Step -1
        Retorno = EnviarDadosGoogle(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, IP_Atualizado, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value)
        If InStr(Retorno, "Sua resposta foi registrada") > 0 Or InStr(Retorno, "Your response has been recorded") > 0 Then ws.Rows(i).Delete
    Next i
    Set ws = Nothing
    Application.ScreenUpdating = True
End Sub

Public Sub DispararFormsSecundario(ByVal ws As Worksheet, ByVal InfoPC As String)
    If ws Is Nothing Then Exit Sub
    Dim NomeSalvo As String, EmailSalvo As String, IP As String
    NomeSalvo = ws.Range("H1").Value
    EmailSalvo = ws.Range("F1").Value
    IP = MeuIP()
    If VerificarInternet() Then
        Call EnviarDadosGoogle(NomeSalvo, InfoPC, Environ("USERNAME"), IP, Format(Now, "dd/mm/yyyy hh:mm:ss"), EmailSalvo)
    Else
        Call SalvarPendenciaOffline(NomeSalvo, InfoPC, Environ("USERNAME"), "Offline", Format(Now, "dd/mm/yyyy hh:mm:ss"), EmailSalvo)
    End If
End Sub

Private Sub MotorDeBlindagemIntegrado(ByVal CaminhoOrigem As String, ByVal caminhoDestinoFinal As String)
    Dim fso As Object, shellApp As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellApp = CreateObject("Shell.Application")
    
    Dim pastaTemp As String, arquivoZipTemp As String, novoZipTemp As String
    Dim pastaExtracao As String, binarioCaminho As String, binarioOculto As String
    Dim caminhoContentTypes As String, caminhoWorkbookRels As String
    Dim itensParaCopiar As Long, canalZip As Integer, tempoEspera As Long
    
    pastaTemp = Environ("TEMP") & "\Nexcel_Blindagem_" & Format(Now, "yymmddhhmmss")
    If fso.FolderExists(pastaTemp) Then fso.DeleteFolder pastaTemp, True
    fso.CreateFolder pastaTemp
    
    arquivoZipTemp = pastaTemp & "\MatrizOrig.zip"
    novoZipTemp = pastaTemp & "\MatrizBlindada.zip"
    pastaExtracao = pastaTemp & "\Extracted"
    fso.CreateFolder pastaExtracao
    
    fso.CopyFile CaminhoOrigem, arquivoZipTemp
    shellApp.Namespace(CVar(pastaExtracao)).CopyHere shellApp.Namespace(CVar(arquivoZipTemp)).Items, 20
    
    caminhoContentTypes = pastaExtracao & "\[Content_Types].xml"
    caminhoWorkbookRels = pastaExtracao & "\xl\_rels\workbook.xml.rels"
    
    tempoEspera = 0
    Do Until fso.FileExists(caminhoWorkbookRels) And fso.FileExists(caminhoContentTypes)
        DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
        If tempoEspera > 5000 Then Exit Do
    Loop
    
    binarioCaminho = pastaExtracao & "\xl\vbaProject.bin"
    binarioOculto = pastaExtracao & "\xl\fontTableCache.bin"
    
    tempoEspera = 0
    Do Until fso.FileExists(binarioCaminho)
        DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
        If tempoEspera > 5000 Then Exit Do
    Loop
    
    tempoEspera = 0
    Do While IsFileLocked(binarioCaminho)
        DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
        If tempoEspera > 5000 Then Exit Do
    Loop
    
    HackParedeConcreto binarioCaminho
    
    SubstituirTextoXML caminhoContentTypes, "vbaProject.bin", "fontTableCache.bin"
    SubstituirTextoXML caminhoWorkbookRels, "vbaProject.bin", "fontTableCache.bin"
    
    On Error Resume Next
    If Dir(binarioOculto) <> "" Then Kill binarioOculto
    Name binarioCaminho As binarioOculto
    On Error GoTo 0
    
    tempoEspera = 0
    Do Until fso.FileExists(binarioOculto)
        DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
        If tempoEspera > 5000 Then Exit Do
    Loop
    
    canalZip = FreeFile
    Open novoZipTemp For Output As #canalZip
    Print #canalZip, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0);
    Close #canalZip
    
    tempoEspera = 0
    Do Until fso.FileExists(novoZipTemp)
        DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
        If tempoEspera > 5000 Then Exit Do
    Loop
    
    itensParaCopiar = shellApp.Namespace(CVar(pastaExtracao)).Items.Count
    shellApp.Namespace(CVar(novoZipTemp)).CopyHere shellApp.Namespace(CVar(pastaExtracao)).Items, 20
    
    tempoEspera = 0
    Do Until shellApp.Namespace(CVar(novoZipTemp)).Items.Count = itensParaCopiar
        DoEvents: Sleep 200: tempoEspera = tempoEspera + 200
        If tempoEspera > 15000 Then Exit Do
    Loop
    
    Sleep 3000
    
    If fso.FileExists(caminhoDestinoFinal) Then fso.DeleteFile caminhoDestinoFinal, True
    fso.CopyFile novoZipTemp, caminhoDestinoFinal
    
LimpezaFinal:
    On Error Resume Next
    fso.DeleteFolder pastaTemp, True
    fso.DeleteFile CaminhoOrigem, True
    On Error GoTo 0
End Sub

Public Sub ReverterBlindagemClient()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    
    Dim arquivosSelecionados As Variant
    Dim pastaDestino As String
    Dim pastaTemp As String, arquivoZipTemp As String, novoZipTemp As String
    Dim binarioCaminho As String, binarioOculto As String, pastaExtracao As String
    Dim itensParaCopiar As Long, canalZip As Integer, i As Long
    Dim caminhoDestinoFinal As String, nomeBase As String, tempoEspera As Long
    Dim caminhoContentTypes As String, caminhoWorkbookRels As String
    
    arquivosSelecionados = Application.GetOpenFilename("Arquivos Excel (*.xlsm), *.xlsm", , "NEXCEL: Selecione os arquivos para DESBLOQUEAR", , True)
    If TypeName(arquivosSelecionados) = "Boolean" Then Exit Sub
    
    With Application.FileDialog(4)
        .Title = "NEXCEL: Onde deseja salvar o(s) sistema(s) DESBLOQUEADO(S)?"
        .InitialFileName = Environ("USERPROFILE") & "\Downloads\"
        If .Show = -1 Then pastaDestino = .SelectedItems(1) & "\" Else Exit Sub
    End With
    
    Application.ScreenUpdating = False
    
    For i = LBound(arquivosSelecionados) To UBound(arquivosSelecionados)
        pastaTemp = Environ("TEMP") & "\Nexcel_Unpack_" & Format(Now, "yymmddhhmmss") & "_" & i
        If fso.FolderExists(pastaTemp) Then fso.DeleteFolder pastaTemp, True
        fso.CreateFolder pastaTemp
        
        arquivoZipTemp = pastaTemp & "\MatrizBlindada.zip"
        novoZipTemp = pastaTemp & "\MatrizDesbloqueada.zip"
        pastaExtracao = pastaTemp & "\Extracted"
        fso.CreateFolder pastaExtracao
        
        fso.CopyFile arquivosSelecionados(i), arquivoZipTemp
        shellApp.Namespace(CVar(pastaExtracao)).CopyHere shellApp.Namespace(CVar(arquivoZipTemp)).Items, 20
        
        binarioCaminho = pastaExtracao & "\xl\vbaProject.bin"
        binarioOculto = pastaExtracao & "\xl\fontTableCache.bin"
        caminhoContentTypes = pastaExtracao & "\[Content_Types].xml"
        caminhoWorkbookRels = pastaExtracao & "\xl\_rels\workbook.xml.rels"
        
        tempoEspera = 0
        Do Until fso.FileExists(caminhoWorkbookRels) And fso.FileExists(caminhoContentTypes)
            DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
            If tempoEspera > 5000 Then Exit Do
        Loop
        
        If fso.FileExists(binarioOculto) And fso.FileExists(caminhoWorkbookRels) And fso.FileExists(caminhoContentTypes) Then
            SubstituirTextoXML caminhoContentTypes, "fontTableCache.bin", "vbaProject.bin"
            SubstituirTextoXML caminhoWorkbookRels, "fontTableCache.bin", "vbaProject.bin"
            
            On Error Resume Next
            If Dir(binarioCaminho) <> "" Then Kill binarioCaminho
            Name binarioOculto As binarioCaminho
            On Error GoTo 0
        End If
        
        tempoEspera = 0
        Do Until fso.FileExists(binarioCaminho)
            DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
            If tempoEspera > 5000 Then Exit Do
        Loop
        
        tempoEspera = 0
        Do While IsFileLocked(binarioCaminho)
            DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
            If tempoEspera > 5000 Then Exit Do
        Loop
        
        HackDerrubarParedeConcreto binarioCaminho
        
        canalZip = FreeFile
        Open novoZipTemp For Output As #canalZip
        Print #canalZip, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0);
        Close #canalZip
        
        tempoEspera = 0
        Do Until fso.FileExists(novoZipTemp)
            DoEvents: Sleep 50: tempoEspera = tempoEspera + 50
            If tempoEspera > 5000 Then Exit Do
        Loop
        
        itensParaCopiar = shellApp.Namespace(CVar(pastaExtracao)).Items.Count
        shellApp.Namespace(CVar(novoZipTemp)).CopyHere shellApp.Namespace(CVar(pastaExtracao)).Items, 20
        
        tempoEspera = 0
        Do Until shellApp.Namespace(CVar(novoZipTemp)).Items.Count = itensParaCopiar
            DoEvents: Sleep 200: tempoEspera = tempoEspera + 200
            If tempoEspera > 15000 Then Exit Do
        Loop
        Sleep 3000
        
        nomeBase = Replace(fso.GetBaseName(arquivosSelecionados(i)), " (Blindado)", "")
        caminhoDestinoFinal = pastaDestino & nomeBase & " (Desbloqueado).xlsm"
        If fso.FileExists(caminhoDestinoFinal) Then fso.DeleteFile caminhoDestinoFinal, True
        fso.CopyFile novoZipTemp, caminhoDestinoFinal
        
        On Error Resume Next
        fso.DeleteFolder pastaTemp, True
        On Error GoTo 0
    Next i
    Application.ScreenUpdating = True
    MsgBox "SUCESSO! Arquivos DESBLOQUEADOS em:" & vbCrLf & pastaDestino, vbInformation, "Nexcel - Reversão"
End Sub

Private Sub SubstituirTextoXML(ByVal CaminhoArquivo As String, ByVal textoAntigo As String, ByVal textoNovo As String)
    Dim objStreamUTF8 As Object
    Dim objStreamBin As Object
    Dim conteudo As String

    Set objStreamUTF8 = CreateObject("ADODB.Stream")
    objStreamUTF8.Type = 2
    objStreamUTF8.Charset = "utf-8"
    objStreamUTF8.Open
    objStreamUTF8.LoadFromFile CaminhoArquivo
    conteudo = objStreamUTF8.ReadText
    objStreamUTF8.Close

    If InStr(1, conteudo, textoAntigo, vbTextCompare) > 0 Then
        conteudo = Replace(conteudo, textoAntigo, textoNovo, 1, -1, vbTextCompare)
        objStreamUTF8.Open
        objStreamUTF8.WriteText conteudo
        objStreamUTF8.Position = 3
        Set objStreamBin = CreateObject("ADODB.Stream")
        objStreamBin.Type = 1
        objStreamBin.Open
        objStreamUTF8.CopyTo objStreamBin
        objStreamBin.SaveToFile CaminhoArquivo, 2
        objStreamBin.Close
        Set objStreamBin = Nothing
        objStreamUTF8.Close
    End If
    Set objStreamUTF8 = Nothing
End Sub

Private Function HackParedeConcreto(caminhoBinario As String) As Boolean
    Dim canal As Integer, b() As Byte, i As Long, j As Long, alterouChave As Boolean
    alterouChave = False
    canal = FreeFile
    Open caminhoBinario For Binary Access Read Write As #canal
        If LOF(canal) > 0 Then
            ReDim b(0 To LOF(canal) - 1)
            Get #canal, 1, b
            For i = UBound(b) - 5 To 0 Step -1
                If b(i) = 68 And b(i + 1) = 80 And b(i + 2) = 66 And b(i + 3) = 61 And b(i + 4) = 34 Then
                    j = i + 5
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 48 To 56: b(j) = b(j) + 1
                            Case 57: b(j) = 48
                            Case 65 To 69: b(j) = b(j) + 1
                            Case 70: b(j) = 65
                            Case 97 To 101: b(j) = b(j) + 1
                            Case 102: b(j) = 97
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 67 And b(i + 1) = 77 And b(i + 2) = 71 And b(i + 3) = 61 And b(i + 4) = 34 Then
                    j = i + 5
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 48 To 56: b(j) = b(j) + 1
                            Case 57: b(j) = 48
                            Case 65 To 69: b(j) = b(j) + 1
                            Case 70: b(j) = 65
                            Case 97 To 101: b(j) = b(j) + 1
                            Case 102: b(j) = 97
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 71 And b(i + 1) = 67 And b(i + 2) = 61 And b(i + 3) = 34 Then
                    j = i + 4
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 48 To 56: b(j) = b(j) + 1
                            Case 57: b(j) = 48
                            Case 65 To 69: b(j) = b(j) + 1
                            Case 70: b(j) = 65
                            Case 97 To 101: b(j) = b(j) + 1
                            Case 102: b(j) = 97
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 73 And b(i + 1) = 68 And b(i + 2) = 61 And b(i + 3) = 34 Then
                    j = i + 4
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 48 To 56: b(j) = b(j) + 1
                            Case 57: b(j) = 48
                            Case 65 To 69: b(j) = b(j) + 1
                            Case 70: b(j) = 65
                            Case 97 To 101: b(j) = b(j) + 1
                            Case 102: b(j) = 97
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
            Next i
            If alterouChave Then Put #canal, 1, b: HackParedeConcreto = True Else HackParedeConcreto = False
        End If
    Close #canal
End Function

Private Function HackDerrubarParedeConcreto(caminhoBinario As String) As Boolean
    Dim canal As Integer, b() As Byte, i As Long, j As Long, alterouChave As Boolean
    alterouChave = False
    canal = FreeFile
    Open caminhoBinario For Binary Access Read Write As #canal
        If LOF(canal) > 0 Then
            ReDim b(0 To LOF(canal) - 1)
            Get #canal, 1, b
            For i = UBound(b) - 5 To 0 Step -1
                If b(i) = 68 And b(i + 1) = 80 And b(i + 2) = 66 And b(i + 3) = 61 And b(i + 4) = 34 Then
                    j = i + 5
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 49 To 57: b(j) = b(j) - 1
                            Case 48: b(j) = 57
                            Case 66 To 70: b(j) = b(j) - 1
                            Case 65: b(j) = 70
                            Case 98 To 102: b(j) = b(j) - 1
                            Case 97: b(j) = 102
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 67 And b(i + 1) = 77 And b(i + 2) = 71 And b(i + 3) = 61 And b(i + 4) = 34 Then
                    j = i + 5
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 49 To 57: b(j) = b(j) - 1
                            Case 48: b(j) = 57
                            Case 66 To 70: b(j) = b(j) - 1
                            Case 65: b(j) = 70
                            Case 98 To 102: b(j) = b(j) - 1
                            Case 97: b(j) = 102
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 71 And b(i + 1) = 67 And b(i + 2) = 61 And b(i + 3) = 34 Then
                    j = i + 4
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 49 To 57: b(j) = b(j) - 1
                            Case 48: b(j) = 57
                            Case 66 To 70: b(j) = b(j) - 1
                            Case 65: b(j) = 70
                            Case 98 To 102: b(j) = b(j) - 1
                            Case 97: b(j) = 102
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
                If b(i) = 73 And b(i + 1) = 68 And b(i + 2) = 61 And b(i + 3) = 34 Then
                    j = i + 4
                    Do While b(j) <> 34 And j < UBound(b)
                        Select Case b(j)
                            Case 49 To 57: b(j) = b(j) - 1
                            Case 48: b(j) = 57
                            Case 66 To 70: b(j) = b(j) - 1
                            Case 65: b(j) = 70
                            Case 98 To 102: b(j) = b(j) - 1
                            Case 97: b(j) = 102
                        End Select
                        j = j + 1
                    Loop
                    alterouChave = True
                End If
            Next i
            If alterouChave Then Put #canal, 1, b: HackDerrubarParedeConcreto = True Else HackDerrubarParedeConcreto = False
        End If
    Close #canal
End Function

Private Function IsFileLocked(CaminhoArquivo As String) As Boolean
    On Error Resume Next
    Dim canal As Integer
    canal = FreeFile
    Open CaminhoArquivo For Input Lock Read As #canal
    If Err.Number <> 0 Then
        IsFileLocked = True
    Else
        IsFileLocked = False
        Close #canal
    End If
    On Error GoTo 0
End Function

Public Sub INJETAR_PROTECAO_VISIBILIDADE(ByVal CaminhoArquivo As String)
    Dim wbDestino As Workbook
    Dim objComponente As Object
    Dim objModulo As Object
    Dim objTwb As Object
    Dim linhaAtual As Long
    Dim qtdLinhasDecl As Long
    Dim achouPrivate As Boolean
    Dim temAppMode As Boolean
    Dim eventosOriginais As Boolean
    Dim strAppMode As String
    
    On Error GoTo TratarErroAcesso
    eventosOriginais = Application.EnableEvents
    Application.EnableEvents = False
    
    Set wbDestino = Workbooks.Open(Filename:=CaminhoArquivo, UpdateLinks:=False, ReadOnly:=False)
    
    If wbDestino.VBProject.Protection = 1 Then GoTo FecharArquivo
    
    For Each objComponente In wbDestino.VBProject.VBComponents
        If objComponente.Type = 1 Then
            Set objModulo = objComponente.CodeModule
            achouPrivate = False
            qtdLinhasDecl = objModulo.CountOfDeclarationLines
            If qtdLinhasDecl > 0 Then
                For linhaAtual = 1 To qtdLinhasDecl
                    If InStr(1, objModulo.Lines(linhaAtual, 1), "Option Private Module", vbTextCompare) > 0 Then
                        achouPrivate = True
                        Exit For
                    End If
                Next linhaAtual
            End If
            If Not achouPrivate Then objModulo.InsertLines 1, "Option Private Module"
        End If
        
        If objComponente.Name = wbDestino.CodeName Then
            Set objTwb = objComponente.CodeModule
            temAppMode = False
            If objTwb.CountOfLines > 0 Then
                If InStr(1, objTwb.Lines(1, objTwb.CountOfLines), "Workbook_Activate", vbTextCompare) > 0 Then temAppMode = True
            End If
            
            If Not temAppMode Then
                strAppMode = ""
                strAppMode = strAppMode & "#If VBA7 Then" & vbCrLf
                strAppMode = strAppMode & "    Private Declare PtrSafe Function FindWindow Lib ""user32"" Alias ""FindWindowA"" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr" & vbCrLf
                strAppMode = strAppMode & "    Private Declare PtrSafe Function GetWindowLong Lib ""user32"" Alias ""GetWindowLongA"" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare PtrSafe Function SetWindowLong Lib ""user32"" Alias ""SetWindowLongA"" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare PtrSafe Function SetWindowPos Lib ""user32"" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare PtrSafe Function GetSystemMetrics Lib ""user32"" (ByVal nIndex As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "#Else" & vbCrLf
                strAppMode = strAppMode & "    Private Declare Function FindWindow Lib ""user32"" Alias ""FindWindowA"" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare Function GetWindowLong Lib ""user32"" Alias ""GetWindowLongA"" (ByVal hwnd As Long, ByVal nIndex As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare Function SetWindowLong Lib ""user32"" Alias ""SetWindowLongA"" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare Function SetWindowPos Lib ""user32"" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "    Private Declare Function GetSystemMetrics Lib ""user32"" (ByVal nIndex As Long) As Long" & vbCrLf
                strAppMode = strAppMode & "#End If" & vbCrLf & vbCrLf
                strAppMode = strAppMode & "Private Const GWL_STYLE As Long = -16" & vbCrLf
                strAppMode = strAppMode & "Private Const WS_CAPTION As Long = &HC00000" & vbCrLf
                strAppMode = strAppMode & "Private Const WS_THICKFRAME As Long = &H40000" & vbCrLf
                strAppMode = strAppMode & "Private Const SWP_FRAMECHANGED As Long = &H20" & vbCrLf
                strAppMode = strAppMode & "Private Const SWP_SHOWWINDOW As Long = &H40" & vbCrLf & vbCrLf
                strAppMode = strAppMode & "Private lOriginalStyle As Long" & vbCrLf & vbCrLf
                strAppMode = strAppMode & "Private Sub Workbook_Activate()" & vbCrLf
                strAppMode = strAppMode & "    On Error Resume Next" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayFormulaBar = False" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayStatusBar = False" & vbCrLf
                strAppMode = strAppMode & "    Application.ExecuteExcel4Macro ""SHOW.TOOLBAR(""""Ribbon"""",False)""" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayFullScreen = True" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayHeadings = False" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayHorizontalScrollBar = False" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayVerticalScrollBar = False" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayWorkbookTabs = True" & vbCrLf
                strAppMode = strAppMode & "    #If VBA7 Then" & vbCrLf
                strAppMode = strAppMode & "        Dim hwnd As LongPtr" & vbCrLf
                strAppMode = strAppMode & "    #Else" & vbCrLf
                strAppMode = strAppMode & "        Dim hwnd As Long" & vbCrLf
                strAppMode = strAppMode & "    #End If" & vbCrLf
                strAppMode = strAppMode & "    hwnd = FindWindow(""XLMAIN"", Application.Caption)" & vbCrLf
                strAppMode = strAppMode & "    If lOriginalStyle = 0 Then lOriginalStyle = GetWindowLong(hwnd, GWL_STYLE)" & vbCrLf
                strAppMode = strAppMode & "    Dim lNewStyle As Long" & vbCrLf
                strAppMode = strAppMode & "    lNewStyle = lOriginalStyle And Not WS_CAPTION And Not WS_THICKFRAME" & vbCrLf
                strAppMode = strAppMode & "    SetWindowLong hwnd, GWL_STYLE, lNewStyle" & vbCrLf
                strAppMode = strAppMode & "    Dim w As Long, h As Long" & vbCrLf
                strAppMode = strAppMode & "    w = GetSystemMetrics(0)" & vbCrLf
                strAppMode = strAppMode & "    h = GetSystemMetrics(1)" & vbCrLf
                strAppMode = strAppMode & "    Application.WindowState = xlNormal" & vbCrLf
                strAppMode = strAppMode & "    SetWindowPos hwnd, 0, 0, 0, w, h, SWP_FRAMECHANGED Or SWP_SHOWWINDOW" & vbCrLf
                strAppMode = strAppMode & "    On Error GoTo 0" & vbCrLf
                strAppMode = strAppMode & "End Sub" & vbCrLf & vbCrLf
                strAppMode = strAppMode & "Private Sub Workbook_Deactivate()" & vbCrLf
                strAppMode = strAppMode & "    On Error Resume Next" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayFormulaBar = True" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayStatusBar = True" & vbCrLf
                strAppMode = strAppMode & "    Application.ExecuteExcel4Macro ""SHOW.TOOLBAR(""""Ribbon"""",True)""" & vbCrLf
                strAppMode = strAppMode & "    Application.DisplayFullScreen = False" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayHeadings = True" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayHorizontalScrollBar = True" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayVerticalScrollBar = True" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayWorkbookTabs = True" & vbCrLf
                strAppMode = strAppMode & "    #If VBA7 Then" & vbCrLf
                strAppMode = strAppMode & "        Dim hwnd As LongPtr" & vbCrLf
                strAppMode = strAppMode & "    #Else" & vbCrLf
                strAppMode = strAppMode & "        Dim hwnd As Long" & vbCrLf
                strAppMode = strAppMode & "    #End If" & vbCrLf
                strAppMode = strAppMode & "    hwnd = FindWindow(""XLMAIN"", Application.Caption)" & vbCrLf
                strAppMode = strAppMode & "    If lOriginalStyle <> 0 Then SetWindowLong hwnd, GWL_STYLE, lOriginalStyle" & vbCrLf
                strAppMode = strAppMode & "    Application.WindowState = xlMaximized" & vbCrLf
                strAppMode = strAppMode & "    On Error GoTo 0" & vbCrLf
                strAppMode = strAppMode & "End Sub" & vbCrLf & vbCrLf
                strAppMode = strAppMode & "Private Sub Workbook_SheetActivate(ByVal Sh As Object)" & vbCrLf
                strAppMode = strAppMode & "    On Error Resume Next" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayHeadings = False" & vbCrLf
                strAppMode = strAppMode & "    ActiveWindow.DisplayGridlines = False" & vbCrLf
                strAppMode = strAppMode & "    On Error GoTo 0" & vbCrLf
                strAppMode = strAppMode & "End Sub"
                objTwb.AddFromString strAppMode
            End If
        End If
    Next objComponente

FecharArquivo:
    wbDestino.Save
    wbDestino.Close SaveChanges:=False
    Application.EnableEvents = eventosOriginais
    Exit Sub
TratarErroAcesso:
    Resume Next
End Sub

Private Sub RestaurarUI_Seguro()
    On Error Resume Next
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.WindowState = xlMaximized
    On Error GoTo 0
End Sub

Public Sub ForcarRenderizacaoCortina()
    On Error Resume Next
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("TRAVADO").Visible = xlSheetVisible
    ThisWorkbook.Sheets("TRAVADO").Activate
    ActiveWindow.DisplayGridlines = False
    DoEvents
    Sleep 150
    On Error GoTo 0
End Sub

Public Sub OcultarAbasProtecaoCTRL()
    Dim ws As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("TRAVADO").Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Sheets
        If UCase(ws.Name) <> "TRAVADO" And UCase(ws.Name) <> UCase(NOME_ABA_OFFLINE) Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub FecharSistemaBlindado()
    Dim wb As Workbook, temOutroVisivel As Boolean
    temOutroVisivel = False
    
    On Error Resume Next
    If Application.Workbooks.Count > 1 Then
        For Each wb In Application.Workbooks
            If UCase(wb.Name) <> UCase(ThisWorkbook.Name) And UCase(wb.Name) <> "PERSONAL.XLSB" Then
                If wb.Windows.Count > 0 Then
                    If wb.Windows(1).Visible = True Then
                        temOutroVisivel = True
                        Exit For
                    End If
                End If
            End If
        Next wb
    End If
    
    If temOutroVisivel Then
        Application.EnableEvents = True
        wb.Activate
        DoEvents
        Application.EnableEvents = False
        ThisWorkbook.Close SaveChanges:=False
        Application.EnableEvents = True
    Else
        Call RestaurarUI_Seguro
        Application.EnableEvents = False
        ThisWorkbook.Saved = True
        Application.Quit
    End If
    On Error GoTo 0
End Sub

Public Sub SAIR_DO_SISTEMA()
    On Error Resume Next
    Dim resposta As Integer
    Dim precisaSalvar As Boolean
    Dim StatusNuvem As String
    
    precisaSalvar = Not ThisWorkbook.Saved
    
    If precisaSalvar Then
        resposta = MsgBox("Deseja salvar as alterações feitas em '" & ThisWorkbook.Name & "'?", vbYesNoCancel + vbQuestion, "Salvar Alterações")
        If resposta = vbCancel Then Exit Sub
    Else
        resposta = vbYes
    End If
    
    Call VerificarPendenciasEnvio
    StatusNuvem = VerificarStatusOnlineCSV()
    
    If StatusNuvem = "BLOQUEADO" Then
        Call Cortina_De_Ferro("SISTEMA BLOQUEADO VIA NUVEM.")
        Exit Sub
    End If
    
    Call OcultarAbasProtecaoCTRL
    Call ForcarRenderizacaoCortina
    
    If precisaSalvar And resposta = vbYes Then
        Application.EnableEvents = False
        ThisWorkbook.Save
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        ThisWorkbook.Saved = True
        Application.EnableEvents = True
    End If
    
    Call FecharSistemaBlindado
    On Error GoTo 0
End Sub
