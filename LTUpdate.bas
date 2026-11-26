Attribute VB_Name = "LTUpdate"
'==============================================================================
'  LT TOOLS - ATUALIZADOR
'  O botão "Atualizar" baixa a versão nova do GitHub
'  (github.com/Linetape/plugin_corel_lttools) para a pasta do LT Tools
'  e troca o código no CorelDRAW. Sem internet, usa os arquivos da pasta.
'==============================================================================
Option Explicit

Private Const T As String = "LT Tools - Atualizar"
Private Const GITHUB_RAW As String = "https://raw.githubusercontent.com/Linetape/plugin_corel_lttools/main/"

'------------------------------------------------------------------------------
'  GitHub
'------------------------------------------------------------------------------
Private Function BaixarTexto(ByVal url As String, ByRef txt As String) As Boolean
    Dim h As Object
    On Error GoTo Falha
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.Open "GET", url & "?t=" & Format$(Now, "yyyymmddhhnnss"), False
    h.setRequestHeader "Cache-Control", "no-cache"
    h.send
    If h.Status <> 200 Then GoTo Falha
    txt = h.responseText
    BaixarTexto = True
    Exit Function
Falha:
    BaixarTexto = False
End Function

Private Function BaixarArquivo(ByVal url As String, ByVal destino As String) As Boolean
    Dim h As Object, st As Object
    On Error GoTo Falha
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.Open "GET", url & "?t=" & Format$(Now, "yyyymmddhhnnss"), False
    h.setRequestHeader "Cache-Control", "no-cache"
    h.send
    If h.Status <> 200 Then GoTo Falha
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                 ' binário
    st.Open
    st.Write h.responseBody
    st.SaveToFile destino, 2    ' sobrescrever
    st.Close
    BaixarArquivo = True
    Exit Function
Falha:
    BaixarArquivo = False
End Function

' versão "2.10" > "2.9"
Private Function VersaoMaior(ByVal a As String, ByVal b As String) As Boolean
    Dim pa As Variant, pb As Variant, i As Long, x As Long, y As Long
    pa = Split(Trim$(a), "."): pb = Split(Trim$(b), ".")
    For i = 0 To 3
        x = 0: y = 0
        If i <= UBound(pa) Then x = Val(pa(i))
        If i <= UBound(pb) Then y = Val(pb(i))
        If x <> y Then VersaoMaior = (x > y): Exit Function
    Next
End Function

' baixa todos os arquivos listados em arquivos.txt para a pasta
Private Function BaixarDoGitHub(ByVal pasta As String, ByRef erro As String) As Boolean
    Dim lista As String, linhas As Variant, i As Long, nome As String, n As Long
    If Not BaixarTexto(GITHUB_RAW & "arquivos.txt", lista) Then erro = "não consegui ler a lista de arquivos": Exit Function
    On Error Resume Next
    MkDir pasta
    MkDir pasta & "icones"
    On Error GoTo 0
    linhas = Split(Replace(lista, vbCr, ""), vbLf)
    For i = 0 To UBound(linhas)
        nome = Trim$(linhas(i))
        If Len(nome) > 0 And Left$(nome, 1) <> "#" Then
            If Not BaixarArquivo(GITHUB_RAW & nome, pasta & Replace(nome, "/", "\")) Then
                erro = "falhou ao baixar " & nome
                Exit Function
            End If
            n = n + 1
        End If
    Next
    BaixarDoGitHub = (n > 0)
End Function

Public Sub LT_Atualizar()
    Dim pasta As String, arqBas As String, arqFrm As String
    Dim app As Object, vbe As Object, proj As Object, comp As Object, p As Object
    Dim txt As String, vNova As String, vAtual As String, passo As String
    Dim vGit As String, erro As String, origem As String

    vAtual = VersaoInstalada()
    pasta = PastaLT()
    If Len(pasta) = 0 Then pasta = Environ$("USERPROFILE") & "\Documents\Corel\Corel Content\LTTools\"

    ' 1) procura versão nova no GitHub
    If BaixarTexto(GITHUB_RAW & "versao.txt", vGit) Then
        vGit = Trim$(Replace(Replace(vGit, vbCr, ""), vbLf, ""))
        If VersaoMaior(vGit, vAtual) Then
            If MsgBox("Tem versão nova do LT Tools no GitHub!" & vbCr & vbCr & _
                      "Versão instalada: " & IIf(Len(vAtual) > 0, vAtual, "?") & vbCr & _
                      "Versão no GitHub: " & vGit & vbCr & vbCr & _
                      "Baixar e instalar agora?", vbQuestion + vbOKCancel, T) <> vbOK Then Exit Sub
            If Not BaixarDoGitHub(pasta, erro) Then
                MsgBox "Não consegui baixar do GitHub (" & erro & ")." & vbCr & _
                       "Tente de novo mais tarde.", vbExclamation, T
                Exit Sub
            End If
            origem = "github"
        Else
            If MsgBox("O LT Tools já está na versão mais nova (" & vAtual & ")." & vbCr & vbCr & _
                      "Reinstalar mesmo assim a partir do GitHub?", vbInformation + vbYesNo + vbDefaultButton2, T) <> vbYes Then Exit Sub
            If Not BaixarDoGitHub(pasta, erro) Then
                MsgBox "Não consegui baixar do GitHub (" & erro & ").", vbExclamation, T
                Exit Sub
            End If
            origem = "github"
        End If
    End If

    If Len(Dir$(pasta & "LTTools.bas")) = 0 Then
        MsgBox "Não consegui acessar o GitHub e não encontrei o LTTools.bas em:" & vbCr & pasta, vbExclamation, T
        Exit Sub
    End If
    arqBas = pasta & "LTTools.bas"
    arqFrm = pasta & "frmLT_codigo.txt"
    vNova = VersaoNoArquivo(arqBas)

    ' 2) sem GitHub: atualiza a partir da pasta
    If origem <> "github" Then
        If MsgBox("Sem acesso ao GitHub agora." & vbCr & _
                  "Atualizar o LT Tools a partir da pasta?" & vbCr & pasta & vbCr & vbCr & _
                  "Versão instalada: " & IIf(Len(vAtual) > 0, vAtual, "?") & vbCr & _
                  "Versão na pasta:  " & IIf(Len(vNova) > 0, vNova, "?"), vbQuestion + vbOKCancel, T) <> vbOK Then Exit Sub
    End If

    On Error GoTo Erro
    passo = "acessar o editor de VBA"
    Set app = Application
    Set vbe = app.VBE
    Set proj = ProjetoDoPlugin(vbe)
    If proj Is Nothing Then Err.Raise 1000, , "Projeto do LT Tools não encontrado."

    ' 1) módulo LTTools: renomeia o antigo, remove e importa o novo
    passo = "trocar o módulo LTTools"
    Set comp = Nothing
    On Error Resume Next
    Set comp = proj.VBComponents("LTTools")
    On Error GoTo Erro
    If Not comp Is Nothing Then
        comp.Name = "LTTools_antigo"
        proj.VBComponents.Remove comp
    End If
    proj.VBComponents.Import arqBas

    ' 2) formulário frmLT: cria se não existir e troca o código
    If Len(Dir$(arqFrm)) > 0 Then
        passo = "atualizar o formulário frmLT"
        Set comp = Nothing
        On Error Resume Next
        Set comp = proj.VBComponents("frmLT")
        On Error GoTo Erro
        If comp Is Nothing Then
            Set comp = proj.VBComponents.Add(3)      ' 3 = UserForm
            comp.Name = "frmLT"
        End If
        txt = LerArquivo(arqFrm)
        With comp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString txt
        End With
    End If

    ' 3) ícones da barra reaplicados sempre que o CorelDRAW abrir
    passo = "configurar a barra ao abrir o CorelDRAW"
    LigarInicio proj

    ' 4) salvar o GlobalMacros
    passo = "salvar o GlobalMacros"
    SalvarProjeto vbe, proj

    MsgBox "LT Tools atualizado" & IIf(Len(vNova) > 0, " para a versão " & vNova, "") & "!" & vbCr & vbCr & "Feche e abra o CorelDRAW para carregar ícones novos da barra.", vbInformation, T
    Exit Sub
Erro:
    MsgBox "Não consegui " & passo & ":" & vbCr & Err.Description & vbCr & vbCr & _
           "Se o erro falar em ""acesso"" ou ""confiança"", faça a troca à mão desta vez" & vbCr & _
           "(remover LTTools, importar LTTools.bas e colar o frmLT_codigo.txt).", vbCritical, T
End Sub

Private Sub LigarInicio(proj As Object)
    Dim comp As Object, cm As Object, txt As String
    On Error Resume Next
    Set comp = proj.VBComponents("ThisMacroStorage")
    If comp Is Nothing Then Exit Sub
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then txt = cm.Lines(1, cm.CountOfLines)
    If InStr(1, txt, "LT_ReaplicarIcones", vbTextCompare) > 0 Then Exit Sub
    If cm.CountOfLines <= 40 Then
        ' só tem código do LT Tools (ou nada): reescreve tudo
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString CodigoInicio()
    Else
        MsgBox "O ThisMacroStorage já tem código seu. Acrescente à mão:" & vbCr & vbCr & CodigoInicio(), vbInformation, T
    End If
End Sub

Private Function CodigoInicio() As String
    CodigoInicio = _
        "Private Sub GlobalMacroStorage_Start()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    LTTools.LT_IniciarBarra" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub GlobalMacroStorage_WindowActivate(ByVal Doc As Document, ByVal Window As Window)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    LTTools.LT_ReaplicarIcones" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub GlobalMacroStorage_DocumentNew(ByVal Doc As Document, ByVal FromTemplate As Boolean, ByVal Template As String, ByVal IncludeGraphics As Boolean)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    LTTools.LT_ReaplicarIcones" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub GlobalMacroStorage_DocumentOpen(ByVal Doc As Document, ByVal FileName As String)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    LTTools.LT_ReaplicarIcones" & vbCrLf & _
        "End Sub"
End Function

Private Sub SalvarProjeto(vbe As Object, proj As Object)
    Dim c As Object
    On Error Resume Next
    Set vbe.ActiveVBProject = proj
    Set c = vbe.CommandBars.FindControl(, 3)     ' 3 = comando "Salvar" do editor
    If Not c Is Nothing Then c.Execute
    If Err.Number <> 0 Or c Is Nothing Then
        On Error GoTo 0
        MsgBox "O código foi atualizado, mas não consegui salvar sozinho." & vbCr & _
               "Abra o editor (Alt+F11) e aperte Ctrl+S.", vbExclamation, T
    End If
End Sub

Private Function PastaLT() As String
    Dim c As Variant, p As String
    p = GetSetting("LTTools", "Config", "Icones", "")
    If Len(p) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        p = Left$(p, InStrRev(Left$(p, Len(p) - 1), "\"))     ' pasta acima de "icones"
        If Len(Dir$(p & "LTTools.bas")) > 0 Then PastaLT = p: Exit Function
    End If
    For Each c In Array(Environ$("USERPROFILE") & "\Documents\Corel\Corel Content\LTTools\", Environ$("USERPROFILE") & "\Documents\LTTools\", "C:\LTTools\", "C:\LTTools\LTTools\", _
                        Environ$("USERPROFILE") & "\Downloads\LTTools\", _
                        Environ$("USERPROFILE") & "\Downloads\LTTools\LTTools\", _
                        Environ$("USERPROFILE") & "\Desktop\LTTools\", _
                        Environ$("USERPROFILE") & "\Desktop\LTTools\LTTools\")
        If Len(Dir$(c & "LTTools.bas")) > 0 Then PastaLT = c: Exit Function
    Next
End Function

Private Function LerArquivo(ByVal arq As String) As String
    Dim f As Integer, s As String
    f = FreeFile
    Open arq For Binary Access Read As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    LerArquivo = s
End Function

Private Function VersaoNoArquivo(ByVal arq As String) As String
    Dim s As String, p As Long, q As Long
    On Error Resume Next
    s = LerArquivo(arq)
    p = InStr(s, "LT_VERSAO As String = """)
    If p > 0 Then
        p = p + Len("LT_VERSAO As String = """)
        q = InStr(p, s, """")
        VersaoNoArquivo = Mid$(s, p, q - p)
    End If
End Function

Private Function VersaoInstalada() As String
    Dim app As Object, c As Object, s As String, p As Long, q As Long, i As Long
    On Error Resume Next
    Set app = Application
    For Each c In app.VBE.VBProjects
        If c.Name = ProjetoDoPlugin(app.VBE).Name Then
            With c.VBComponents("LTTools").CodeModule
                For i = 1 To .CountOfDeclarationLines
                    s = .Lines(i, 1)
                    p = InStr(s, "LT_VERSAO As String = """)
                    If p > 0 Then
                        p = p + Len("LT_VERSAO As String = """)
                        q = InStr(p, s, """")
                        VersaoInstalada = Mid$(s, p, q - p)
                        Exit Function
                    End If
                Next
            End With
        End If
    Next
End Function

' projeto onde este atualizador está: LTPlugin (instalado pelo .exe) ou GlobalMacros
Private Function ProjetoDoPlugin(vbe As Object) As Object
    Dim p As Object, c As Object, alt As Object
    On Error Resume Next
    For Each p In vbe.VBProjects
        Set c = Nothing
        Set c = p.VBComponents("LTUpdate")
        If Not c Is Nothing Then
            If LCase$(p.Name) = "ltplugin" Then Set ProjetoDoPlugin = p: Exit Function
            If alt Is Nothing Then Set alt = p
        End If
    Next
    Set ProjetoDoPlugin = alt
End Function
