Attribute VB_Name = "LTTools"
'==============================================================================
'  LT TOOLS  -  Plugin VBA para CorelDRAW (X7 ou superior)
'  Line Tape Iluminação e Locação - Goiânia/GO
'
'  Conjunto de ferramentas para produção (laser, router CNC, plotter,
'  letra caixa, LED e neon), inspirado nas funções apresentadas no canal
'  "eCut Brasil". Código próprio, escrito do zero - NÃO é o eCut.
'
'  INSTALAÇÃO: veja o arquivo LEIA-ME.txt (pasta C:\LTTools).
'  Depois de instalado, rode a macro LT_InstalarBarra uma vez para criar
'  a barra de ícones "LT Tools".
'
'  Todas as medidas são em MILÍMETROS. Cada ferramenta pode ser desfeita
'  com um único Ctrl+Z.
'==============================================================================
Option Explicit

Private Const TITULO As String = "LT Tools"
Public Const LT_VERSAO As String = "2.5"
Private Const PI As Double = 3.14159265358979
Private Const AMOSTRAS As Long = 12      ' pontos por segmento de curva (precisão)

Private gUnidade As Long
Private gAtivo As Boolean

' variáveis usadas pelo gerador de quebra-cabeça
Private gAx As Double, gAy As Double, gUx As Double, gUy As Double
Private gNx As Double, gNy As Double, gL As Double, gT As Double

' dados trocados com o formulário frmLT (janelas e menus)
Public gModo As Long                 ' 1 = janela de parâmetros, 2 = menu
Public gTit As String, gDica As String
Public gN As Long, gOK As Boolean
Public gTipo(1 To 16) As Long        ' 1 número, 2 sim/não, 3 lista, 4 texto
Public gRot(1 To 16) As String
Public gOpc(1 To 16) As String
Public gVal(1 To 16) As String
Public gMenuN As Long, gMenuEscolha As Long
Public gMenuTxt(1 To 16) As String
Public gMenuIco(1 To 16) As String
Public gFab(1 To 16) As String, gUlt(1 To 16) As String   ' valores de fábrica / último usado
Public gConfig As Boolean

' --- encaixe pela forma (raster) ---
Private nzN As Long
Private nzShp() As Shape
Private nzPX() As Variant, nzPY() As Variant, nzPF() As Variant   ' polígonos por peça
Private nzCX() As Double, nzCY() As Double
Private nzPagW As Double, nzPagH As Double
Private nzSX As Double, nzSY As Double, nzSW As Double, nzSH As Double
' parâmetros do cálculo atual
Private nzSWch As Double, nzSHch As Double, nzMarg As Double, nzRes As Double
Private nzGW As Long, nzGH As Long, nzK As Long, nzNA As Long, nzAng() As Double
Private nzAssin As String
' rasters (peça x ângulo)
Private rsW() As Long, rsH() As Long, rsXm() As Double, rsYm() As Double, rsOk() As Boolean, rsArea() As Long
Private rsR0() As Long, rsR1() As Long, rsD0() As Long, rsD1() As Long, rsKW() As Long
Private ruR() As Long, ruA() As Long, ruB() As Long, nRu As Long
Private rdR() As Long, rdA() As Long, rdB() As Long, nRd As Long
' resultado
Private nzS() As Long, nzAI() As Long, nzX() As Long, nzY() As Long, nzNS As Long, nzFora As Long
Private nzAreaPecas As Double
Public gPreview As Boolean, gPrevArq As String, gPrevInfo As String

Public gBarraOK As Boolean           ' ícones da barra já aplicados nesta sessão            ' True = só configurar o padrão (não executa)

' estilos (rótulos = chaves salvas no registro)
Private Const CORES As String = "Magenta|Vermelho|Azul|Verde|Preto|Ciano|Amarelo|Laranja|Branco|Cinza"
Private Const E_CORTE As String = "Cor da linha de corte"
Private Const E_ESPCORTE As String = "Espessura da linha de corte (mm)"
Private Const E_FURO As String = "Cor dos furos e ilhoses"
Private Const E_MARCA As String = "Cor das marcações, cotas e chapas"
Private Const E_ESPMARCA As String = "Espessura das marcações (mm)"
Private Const E_FONTE As String = "Fonte dos textos"
Private Const E_CORTXT As String = "Cor dos textos"
Private Const E_LED As String = "Cor dos módulos de LED"
Private Const E_CAMADAS As String = "Criar camadas separadas (CORTE, FUROS...)"

'==============================================================================
'  MENUS - cada macro abaixo é um botão da barra "LT Tools"
'==============================================================================
Public Sub LT_Menu()
    Select Case AbrirMenu("LT Tools", _
        Array("Medição e orçamento", "Encaixe e otimização de corte", "Plotter, laser e CNC", _
              "Letra caixa, LED e neon", "Geradores", "Edição de objetos", _
              "Estilos (cores, linhas, fonte)", "Valores padrão das ferramentas", "Instalar / atualizar a barra"), _
        Array("cat_medicao", "cat_encaixe", "cat_corte", "cat_letracaixa", "cat_geradores", "cat_edicao", _
              "estilos", "config", "logo"))
        Case 1: LT_Medicao
        Case 2: LT_Encaixe
        Case 3: LT_Corte
        Case 4: LT_LetraCaixaLED
        Case 5: LT_Geradores
        Case 6: LT_Edicao
        Case 7: LT_Estilos
        Case 8: LT_Configurar
        Case 9: LT_InstalarBarra
    End Select
End Sub

Public Sub LT_Medicao()
    Select Case AbrirMenu("Medição e orçamento", _
        Array("Área, perímetro e nós", "Orçamento (material + corte)", "Tempo de corte", _
              "Volume e massa (peso)", "Centro de massa", "Cotas automáticas"), _
        Array("area", "orcamento", "tempo", "massa", "centro", "cotas"))
        Case 1: LT_AreaPerimetro
        Case 2: LT_Orcamento
        Case 3: LT_TempoCorte
        Case 4: LT_VolumeMassa
        Case 5: LT_CentroMassa
        Case 6: LT_Cotas
    End Select
End Sub

Public Sub LT_Encaixe()
    Select Case AbrirMenu("Encaixe e otimização", _
        Array("Encaixe pela forma (giro livre, com prévia)", "Encaixe retangular em chapas", "Multiplicar / preencher chapa", "Remover duplicados", _
              "Otimizar percurso de corte", "Fechar curvas abertas", "Simplificar curvas (reduzir nós)", _
              "Agrupamento avançado", "Soldar sobrepostos"), _
        Array("enc_forma", "enc_ret", "multiplicar", "duplicados", "percurso", "fechar", "simplificar", "agrupar", "soldar"))
        Case 1: LT_EncaixeForma
        Case 2: LT_EncaixeRetangular
        Case 3: LT_Multiplicar
        Case 4: LT_RemoverDuplicados
        Case 5: LT_OtimizarPercurso
        Case 6: LT_FecharCurvas
        Case 7: LT_SimplificarCurvas
        Case 8: LT_AgrupamentoAvancado
        Case 9: LT_SoldarSobrepostos
    End Select
End Sub

Public Sub LT_Corte()
    Select Case AbrirMenu("Plotter, laser e CNC", _
        Array("Contorno de corte (impressão + recorte)", "Marcas de registro", "Linhas de descarte (vinil)", _
              "Fatiar em painéis", "Gerar G-code (router CNC)", "Exportar DXF"), _
        Array("contorno", "marcas", "descarte", "fatiar", "gcode", "dxf"))
        Case 1: LT_ContornoCorte
        Case 2: LT_MarcasRegistro
        Case 3: LT_LinhasDescarte
        Case 4: LT_Fatiar
        Case 5: LT_GCode
        Case 6: LT_ExportarDXF
    End Select
End Sub

Public Sub LT_LetraCaixaLED()
    Select Case AbrirMenu("Letra caixa, LED e neon", _
        Array("Letra caixa (laterais)", "Distribuir módulos de LED", "Neon", _
              "Furos de montagem / drenagem", "Ilhoses (lona)"), _
        Array("letracaixa", "led", "neon", "furos", "ilhoses"))
        Case 1: LT_LetraCaixa
        Case 2: LT_DistribuirLED
        Case 3: LT_Neon
        Case 4: LT_FurosContorno
        Case 5: LT_Ilhoses
    End Select
End Sub

Public Sub LT_Geradores()
    Select Case AbrirMenu("Geradores", _
        Array("Caixa com malhete", "Quebra-cabeça", "Grelha (grade de peças)", "Círculos aleatórios", _
              "Numerar objetos", "Sequência numerada (rifa, comanda)", "Caixa delimitadora"), _
        Array("malhete", "quebracabeca", "grelha", "circulos", "numerar", "sequencia", "caixadelim"))
        Case 1: LT_CaixaMalhete
        Case 2: LT_QuebraCabeca
        Case 3: LT_Grelha
        Case 4: LT_CirculosAleatorios
        Case 5: LT_NumerarObjetos
        Case 6: LT_SequenciaNumerada
        Case 7: LT_CaixaDelimitadora
    End Select
End Sub

Public Sub LT_Edicao()
    Select Case AbrirMenu("Edição de objetos", _
        Array("Definir tamanho exato", "Distribuir com espaçamento", "Objetos ao longo da curva", _
              "Trocar objetos pelo modelo", "Espelhar por linha guia", "Medir ângulo da linha", _
              "Arredondar cantos", "Separar curvas mantendo furos"), _
        Array("tamanho", "distribuir", "aolongo", "trocar", "espelhar", "angulo", "arredondar", "separar"))
        Case 1: LT_DefinirTamanho
        Case 2: LT_DistribuirEspaco
        Case 3: LT_ObjetosAoLongoCurva
        Case 4: LT_TrocarObjetos
        Case 5: LT_EspelharPorLinha
        Case 6: LT_AnguloLinha
        Case 7: LT_ArredondarCantos
        Case 8: LT_SepararInteligente
    End Select
End Sub

'==============================================================================
'  INSTALAR A BARRA DE FERRAMENTAS
'==============================================================================
Public Sub LT_InstalarBarra()
    InstalarBarraInterno False
End Sub

' chamado automaticamente quando o CorelDRAW abre (ThisMacroStorage)
Public Sub LT_IniciarBarra()
    InstalarBarraInterno True
End Sub

' chamado quando uma janela/documento é ativado: o Corel já terminou de montar as
' barras, então os ícones aplicados aqui não são substituídos (1 vez por sessão)
Public Sub LT_ReaplicarIcones()
    Static feito As Boolean
    If feito Then Exit Sub
    feito = True
    InstalarBarraInterno True
End Sub

Private Sub InstalarBarraInterno(ByVal silencioso As Boolean)
    Const CAT_MACROS As String = "2cc24a3e-fe24-4708-9a74-9c75406eebcd"
    Dim app As Object, cbs As Object, cb As Object, ctl As Object, i As Long, pasta As String
    Dim mac As Variant, ico As Variant, dica As Variant, falhas As String
    pasta = PastaIcones()
    If Dir$(pasta & "logo.png") = "" Then
        If silencioso Then Exit Sub
        pasta = InputBox("Não encontrei os ícones em:" & vbCr & pasta & vbCr & vbCr & _
                         "Informe a pasta onde está a pasta ""icones"" do LT Tools:", TITULO, Environ$("USERPROFILE") & "\Documents\LTTools\icones\")
        If Len(pasta) = 0 Then Exit Sub
        If Right$(pasta, 1) <> "\" Then pasta = pasta & "\"
        SaveSetting "LTTools", "Config", "Icones", pasta
    End If
    mac = Array("LTTools.LT_Menu", "LTTools.LT_Medicao", "LTTools.LT_Encaixe", "LTTools.LT_Corte", _
                "LTTools.LT_LetraCaixaLED", "LTTools.LT_Geradores", "LTTools.LT_Edicao", "LTUpdate.LT_Atualizar")
    ico = Array("logo", "cat_medicao", "cat_encaixe", "cat_corte", "cat_letracaixa", "cat_geradores", "cat_edicao", "atualizar")
    dica = Array("LT Tools " & LT_VERSAO, "Medição e orçamento", "Encaixe e otimização", "Plotter, laser e CNC", _
                 "Letra caixa, LED e neon", "Geradores", "Edição de objetos", "Atualizar LT Tools")
    On Error GoTo Erro
    Set app = Application
    Set cbs = app.CommandBars
    On Error Resume Next
    Set cb = cbs.Item("LT Tools")
    On Error GoTo Erro
    ' ao abrir o Corel (silencioso) só completa a barra e reaplica os ícones,
    ' mantendo o lugar onde ela foi encaixada
    If Not silencioso And Not cb Is Nothing Then
        cb.Delete: Set cb = Nothing
    End If
    If cb Is Nothing Then Set cb = cbs.Add("LT Tools")
    For i = 0 To UBound(mac)
        If cb.Controls.Count > i Then
            Set ctl = cb.Controls.Item(i + 1)
        Else
            Set ctl = cb.Controls.AddCustomButton(CAT_MACROS, ProjetoLT() & "." & mac(i))
        End If
        On Error Resume Next
        ctl.SetIcon2 pasta & ico(i) & ".png"          ' CorelDRAW 2026: SetIcon2 com PNG
        If Err.Number <> 0 Then
            Err.Clear
            ctl.SetCustomIcon pasta & ico(i) & ".ico"
            If Err.Number <> 0 Then falhas = falhas & ico(i) & " ": Err.Clear
        End If
        ctl.Caption = dica(i)
        ctl.ToolTipText = dica(i)
        On Error GoTo Erro
    Next
    cb.Visible = True
    gBarraOK = (Len(falhas) = 0)
    If silencioso Then Exit Sub
    MsgBox "Barra ""LT Tools"" criada!" & vbCr & _
           "Arraste-a para junto das outras barras, no topo da tela." & _
           IIf(Len(falhas) > 0, vbCr & vbCr & "Ícones não carregados: " & falhas, ""), vbInformation, TITULO
    Exit Sub
Erro:
    If silencioso Then Exit Sub
    MsgBox "Não consegui criar a barra automaticamente (" & Err.Description & ")." & vbCr & vbCr & _
           "Faça à mão: Ferramentas > Opções > Personalização > Barras de comando > Novo (""LT Tools"")." & vbCr & _
           "Depois, em Comandos > Scripts, arraste LT_Medicao, LT_Encaixe, LT_Corte," & vbCr & _
           "LT_LetraCaixaLED, LT_Geradores e LT_Edicao para a barra e use" & vbCr & _
           "Aparência > Importar para escolher os ícones da pasta ""icones"".", vbExclamation, TITULO
End Sub
'==============================================================================
'  1. MEDIÇÃO E ORÇAMENTO
'==============================================================================
Public Sub LT_AreaPerimetro()
    Dim a As Double, p As Double, e As Long, nos As Long, q As Long
    If Not Iniciar("Área e perímetro") Then Exit Sub
    On Error GoTo Erro
    Totais a, p, e, nos, q
    Finalizar
    MsgBox "Objetos: " & q & vbCr & _
           "Área: " & Fm(a, 1) & " mm²  =  " & Fm(a / 1000000#, 4) & " m²" & vbCr & _
           "Perímetro (corte): " & Fm(p, 1) & " mm  =  " & Fm(p / 1000#, 3) & " m" & vbCr & _
           "Entradas / furos de início (contornos): " & e & vbCr & _
           "Nós: " & nos, vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_Orcamento()
    Dim a As Double, p As Double, e As Long, nos As Long, q As Long
    Dim pm2 As Double, perda As Double, pml As Double, pent As Double, marg As Double
    Dim cMat As Double, cCorte As Double, cEnt As Double, tot As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Orçamento"
    DlgNum "Preço do material (R$ por m²)", 120
    DlgNum "Perda de material (%)", 15
    DlgNum "Preço do corte (R$ por metro)", 2
    DlgNum "Preço por entrada / furo de início (R$)", 0.1
    DlgNum "Margem / lucro (%)", 30
    If Not DlgMostrar() Then Exit Sub
    pm2 = DV(1): perda = DV(2): pml = DV(3): pent = DV(4)
    marg = DV(5)
    If Not Iniciar("Orçamento") Then Exit Sub
    On Error GoTo Erro
    Totais a, p, e, nos, q
    Finalizar
    cMat = (a / 1000000#) * (1 + perda / 100) * pm2
    cCorte = (p / 1000#) * pml
    cEnt = e * pent
    tot = (cMat + cCorte + cEnt) * (1 + marg / 100)
    MsgBox "ORÇAMENTO - " & q & " objeto(s)" & vbCr & vbCr & _
           "Área: " & Fm(a / 1000000#, 4) & " m² (+" & Fm(perda, 0) & "% perda)" & vbCr & _
           "Material: R$ " & Fm(cMat) & vbCr & _
           "Corte: " & Fm(p / 1000#, 2) & " m  ->  R$ " & Fm(cCorte) & vbCr & _
           "Entradas: " & e & "  ->  R$ " & Fm(cEnt) & vbCr & _
           "Custo: R$ " & Fm(cMat + cCorte + cEnt) & vbCr & _
           "Margem: " & Fm(marg, 0) & "%" & vbCr & vbCr & _
           "PREÇO FINAL: R$ " & Fm(tot), vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_TempoCorte()
    Dim a As Double, p As Double, e As Long, nos As Long, q As Long
    Dim v As Double, tf As Double, pas As Double, seg As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Tempo de corte"
    DlgNum "Velocidade de corte (mm/s)", 20
    DlgNum "Tempo por entrada / perfuração (s)", 0.5
    DlgNum "Número de passadas", 1
    If Not DlgMostrar() Then Exit Sub
    v = DV(1): tf = DV(2): pas = DV(3)
    If v <= 0 Then Exit Sub
    If Not Iniciar("Tempo de corte") Then Exit Sub
    On Error GoTo Erro
    Totais a, p, e, nos, q
    Finalizar
    seg = (p / v + e * tf) * pas
    seg = seg * 1.1   ' +10% deslocamentos em vazio / aceleração
    MsgBox "Comprimento de corte: " & Fm(p / 1000#, 2) & " m" & vbCr & _
           "Entradas: " & e & "   Passadas: " & pas & vbCr & vbCr & _
           "TEMPO ESTIMADO: " & HMS(seg) & vbCr & _
           "(inclui +10% para deslocamentos em vazio)", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_VolumeMassa()
    Dim a As Double, p As Double, e As Long, nos As Long, q As Long
    Dim esp As Double, dens As Double, vol As Double
    Dim mat As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Volume e massa"
    DlgNum "Espessura do material (mm)", 3
    DlgLista "Material", "Acrílico (1,19)|MDF (0,75)|PVC expandido (0,55)|ACM (1,38)|Alumínio (2,70)|Aço (7,85)|Poliestireno (1,05)|Outro (usar densidade abaixo)", 1
    DlgNum "Densidade se for ""Outro"" (g/cm³)", 1
    If Not DlgMostrar() Then Exit Sub
    esp = DV(1): mat = DV(2): dens = DV(3)
    If mat >= 1 And mat <= 7 Then dens = Choose(mat, 1.19, 0.75, 0.55, 1.38, 2.7, 7.85, 1.05)
    If Not Iniciar("Volume e massa") Then Exit Sub
    On Error GoTo Erro
    Totais a, p, e, nos, q
    Finalizar
    vol = a * esp / 1000#                 ' cm³
    MsgBox "Área: " & Fm(a / 1000000#, 4) & " m²" & vbCr & _
           "Volume: " & Fm(vol, 1) & " cm³" & vbCr & _
           "MASSA: " & Fm(vol * dens / 1000#, 3) & " kg", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_CentroMassa()
    Dim col As Collection, s As Shape, cx As Double, cy As Double, n As Long
    Dim lay As Layer, r As Double, x As Shape
    If Not Iniciar("Centro de massa") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("CENTRO DE MASSA")
    Set col = Folhas()
    For Each s In col
        If Centroide(s, cx, cy) Then
            r = 3
            Set x = lay.CreateEllipse2(cx, cy, r)
            Estilo x, "marca"
            Set x = lay.CreateLineSegment(cx - r * 2, cy, cx + r * 2, cy)
            Estilo x, "marca"
            Set x = lay.CreateLineSegment(cx, cy - r * 2, cx, cy + r * 2)
            Estilo x, "marca"
            n = n + 1
        End If
    Next
    Finalizar
    MsgBox n & " centro(s) de massa marcado(s) na camada ""CENTRO DE MASSA"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_Cotas()
    Dim arr() As Shape, n As Long, i As Long, lay As Layer
    Dim x As Double, y As Double, w As Double, h As Double, off As Double, tam As Double
    Dim t As Shape
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Cotas automáticas"
    DlgNum "Distância da cota até o objeto (mm)", 8
    DlgNum "Tamanho do texto (pt)", 10
    If Not DlgMostrar() Then Exit Sub
    off = DV(1): tam = DV(2)
    If Not Iniciar("Cotas") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("COTAS")
    n = Itens(arr)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        ' horizontal
        CotaLinha lay, x, y - off, x + w, y - off
        lay.CreateLineSegment(x, y - off - 2, x, y - 1).Outline.Width = 0.1
        lay.CreateLineSegment(x + w, y - off - 2, x + w, y - 1).Outline.Width = 0.1
        Texto lay, Fm(w, 1) & " mm", x + w / 2, y - off - 2 - tam * 0.25, tam
        ' vertical
        CotaLinha lay, x - off, y, x - off, y + h
        lay.CreateLineSegment(x - off - 2, y, x - 1, y).Outline.Width = 0.1
        lay.CreateLineSegment(x - off - 2, y + h, x - 1, y + h).Outline.Width = 0.1
        Set t = Texto(lay, Fm(h, 1) & " mm", x - off - 2 - tam * 0.25, y + h / 2, tam)
        t.RotateEx 90, x - off - 2 - tam * 0.25, y + h / 2
    Next
    Finalizar
    Exit Sub
Erro:
    Falha
End Sub

Private Sub CotaLinha(lay As Layer, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim s As Shape
    Set s = lay.CreateLineSegment(x1, y1, x2, y2)
    Estilo s, "marca"
End Sub

'==============================================================================
'  2. ENCAIXE E OTIMIZAÇÃO
'==============================================================================
Public Sub LT_EncaixeRetangular()
    ' Encaixe MaxRects: cada peça testa as posições normal e girada 90° em todos os
    ' espaços livres de todas as chapas e fica onde encaixa mais justo.
    Dim arr() As Shape, n As Long, i As Long, k As Long, f As Long, g As Long, t As Long, tMax As Long
    Dim chW As Double, chH As Double, gap As Double, marg As Double, rot As Double
    Dim bw() As Double, bh() As Double, idx() As Long, chave() As Double
    Dim sx As Double, sy As Double, sw As Double, sh As Double
    Dim x As Double, y As Double, w As Double, h As Double, UW As Double, UH As Double
    Dim frX() As Double, frY() As Double, frW() As Double, frH() As Double, frS() As Long, nf As Long
    Dim nCh As Long, fora As Long, girados As Long, areaPecas As Double, r As Shape, lay As Layer
    Dim melhor As Double, sc As Double, bF As Long, bRot As Boolean, pw As Double, ph As Double
    Dim px As Double, py As Double, ox As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Encaixe retangular", "Encaixa pelo retângulo de cada peça. Peças agrupadas são tratadas como uma só."
    DlgNum "Largura da chapa (mm)", 1220
    DlgNum "Altura da chapa (mm)", 2440
    DlgNum "Espaço entre peças (mm)", 3
    DlgNum "Margem da chapa (mm)", 5
    DlgSim "Girar peças 90° quando encaixar melhor", 1
    If Not DlgMostrar() Then Exit Sub
    chW = DV(1): chH = DV(2): gap = DV(3): marg = DV(4): rot = DV(5)
    If Not Iniciar("Encaixe retangular") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    ActiveSelectionRange.GetBoundingBox sx, sy, sw, sh
    UW = chW - 2 * marg + gap: UH = chH - 2 * marg + gap
    ReDim bw(1 To n): ReDim bh(1 To n): ReDim idx(1 To n): ReDim chave(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        bw(i) = w: bh(i) = h: idx(i) = i
        chave(i) = -(w * h)                      ' maiores primeiro
        areaPecas = areaPecas + w * h
    Next
    OrdenarIdx idx, chave, n
    ReDim frX(1 To 16): ReDim frY(1 To 16): ReDim frW(1 To 16): ReDim frH(1 To 16): ReDim frS(1 To 16)
    tMax = IIf(rot = 1, 1, 0)
    Set lay = Camada("CHAPAS")
    For k = 1 To n
        i = idx(k)
        If Not (bw(i) + gap <= UW + 0.001 And bh(i) + gap <= UH + 0.001) And _
           Not (rot = 1 And bh(i) + gap <= UW + 0.001 And bw(i) + gap <= UH + 0.001) Then
            fora = fora + 1                      ' maior que a chapa
        Else
            Do
                melhor = 1E+300: bF = 0
                For f = 1 To nf
                    For t = 0 To tMax
                        If t = 0 Then pw = bw(i) + gap: ph = bh(i) + gap Else pw = bh(i) + gap: ph = bw(i) + gap
                        If pw <= frW(f) + 0.001 And ph <= frH(f) + 0.001 Then
                            sc = frS(f) * 1E+12 + Min2(frW(f) - pw, frH(f) - ph) * 1000000# + Max2(frW(f) - pw, frH(f) - ph)
                            If sc < melhor Then melhor = sc: bF = f: bRot = (t = 1)
                        End If
                    Next
                Next
                If bF > 0 Then Exit Do
                nCh = nCh + 1                    ' não coube: abre outra chapa
                AddLivre frX, frY, frW, frH, frS, nf, 0, 0, UW, UH, nCh
                ox = sx + sw + 50 + (nCh - 1) * (chW + 50)
                Set r = lay.CreateRectangle2(ox, sy, chW, chH)
                r.Fill.ApplyNoFill
                Estilo r, "marca"
            Loop
            If bRot Then pw = bh(i) + gap: ph = bw(i) + gap Else pw = bw(i) + gap: ph = bh(i) + gap
            px = frX(bF): py = frY(bF): g = frS(bF)
            If bRot Then arr(i).Rotate 90: girados = girados + 1
            arr(i).GetBoundingBox x, y, w, h
            ox = sx + sw + 50 + (g - 1) * (chW + 50)
            arr(i).Move ox + marg + px - x, sy + marg + py - y
            DividirLivres frX, frY, frW, frH, frS, nf, px, py, pw, ph, g
        End If
    Next
    Finalizar
    If nCh = 0 Then
        MsgBox "Nenhuma peça cabe na chapa informada.", vbExclamation, TITULO
    Else
        MsgBox "Chapas usadas: " & nCh & vbCr & _
               "Aproveitamento: " & Fm(100 * areaPecas / (nCh * chW * chH), 1) & "%" & vbCr & _
               "Peças giradas: " & girados & vbCr & _
               IIf(fora > 0, fora & " peça(s) maiores que a chapa ficaram no lugar.", ""), vbInformation, TITULO
    End If
    Exit Sub
Erro:
    Falha
End Sub

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Sub AddLivre(frX() As Double, frY() As Double, frW() As Double, frH() As Double, frS() As Long, nf As Long, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal s As Long)
    If w <= 0.001 Or h <= 0.001 Then Exit Sub
    nf = nf + 1
    If nf > UBound(frX) Then
        ReDim Preserve frX(1 To nf * 2): ReDim Preserve frY(1 To nf * 2): ReDim Preserve frW(1 To nf * 2)
        ReDim Preserve frH(1 To nf * 2): ReDim Preserve frS(1 To nf * 2)
    End If
    frX(nf) = x: frY(nf) = y: frW(nf) = w: frH(nf) = h: frS(nf) = s
End Sub

Private Sub TirarLivre(frX() As Double, frY() As Double, frW() As Double, frH() As Double, frS() As Long, nf As Long, ByVal k As Long)
    frX(k) = frX(nf): frY(k) = frY(nf): frW(k) = frW(nf): frH(k) = frH(nf): frS(k) = frS(nf)
    nf = nf - 1
End Sub

' recorta os espaços livres pela peça colocada e remove os que ficaram dentro de outros
Private Sub DividirLivres(frX() As Double, frY() As Double, frW() As Double, frH() As Double, frS() As Long, nf As Long, _
                          ByVal px As Double, ByVal py As Double, ByVal pw As Double, ByVal ph As Double, ByVal s As Long)
    Dim f As Long, a As Long, b As Long, fx As Double, fy As Double, fw As Double, fh As Double, dentro As Boolean
    f = 1
    Do While f <= nf
        If frS(f) = s And px < frX(f) + frW(f) - 0.001 And px + pw > frX(f) + 0.001 And _
           py < frY(f) + frH(f) - 0.001 And py + ph > frY(f) + 0.001 Then
            fx = frX(f): fy = frY(f): fw = frW(f): fh = frH(f)
            TirarLivre frX, frY, frW, frH, frS, nf, f
            If px > fx Then AddLivre frX, frY, frW, frH, frS, nf, fx, fy, px - fx, fh, s
            If px + pw < fx + fw Then AddLivre frX, frY, frW, frH, frS, nf, px + pw, fy, fx + fw - px - pw, fh, s
            If py > fy Then AddLivre frX, frY, frW, frH, frS, nf, fx, fy, fw, py - fy, s
            If py + ph < fy + fh Then AddLivre frX, frY, frW, frH, frS, nf, fx, py + ph, fw, fy + fh - py - ph, s
        Else
            f = f + 1
        End If
    Loop
    a = 1
    Do While a <= nf
        dentro = False
        For b = 1 To nf
            If a <> b And frS(a) = frS(b) Then
                If frX(a) >= frX(b) - 0.001 And frY(a) >= frY(b) - 0.001 And _
                   frX(a) + frW(a) <= frX(b) + frW(b) + 0.001 And frY(a) + frH(a) <= frY(b) + frH(b) + 0.001 Then
                    dentro = True: Exit For
                End If
            End If
        Next
        If dentro Then TirarLivre frX, frY, frW, frH, frS, nf, a Else a = a + 1
    Loop
End Sub

Public Sub LT_Multiplicar()
    Dim cols As Double, lins As Double, gap As Double, aW As Double, aH As Double
    Dim sr As ShapeRange, d As ShapeRange, x As Double, y As Double, w As Double, h As Double
    Dim c As Long, l As Long, tot As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Multiplicar"
    DlgLista "Modo", "Colunas x linhas|Preencher uma área / chapa", 1
    DlgNum "Colunas", 3
    DlgNum "Linhas", 2
    DlgNum "Largura da área (modo preencher, mm)", 1000
    DlgNum "Altura da área (modo preencher, mm)", 1000
    DlgNum "Espaço entre cópias (mm)", 5
    If Not DlgMostrar() Then Exit Sub
    cols = DV(2): lins = DV(3): aW = DV(4): aH = DV(5): gap = DV(6)
    If DV(1) = 2 Then cols = 0
    If Not Iniciar("Multiplicar") Then Exit Sub
    On Error GoTo Erro
    Set sr = ActiveSelectionRange
    sr.GetBoundingBox x, y, w, h
    If cols = 0 Then
        cols = Int((aW + gap) / (w + gap))
        lins = Int((aH + gap) / (h + gap))
    End If
    For l = 0 To lins - 1
        For c = 0 To cols - 1
            If c > 0 Or l > 0 Then
                Set d = sr.Duplicate
                d.Move c * (w + gap), -l * (h + gap)
            End If
            tot = tot + 1
        Next
    Next
    Finalizar
    MsgBox "Total: " & tot & " unidade(s) (" & cols & " x " & lins & ").", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_RemoverDuplicados()
    Dim col As Collection, s As Shape, n As Long, i As Long, j As Long, tol As Double
    Dim fx() As Double, fy() As Double, fw() As Double, fh() As Double, fl() As Double, fn() As Long
    Dim apagar() As Boolean, arr() As Shape, c As Object, so As Object, del As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Remover duplicados"
    DlgNum "Tolerância (mm)", 0.1
    If Not DlgMostrar() Then Exit Sub
    tol = DV(1)
    If Not Iniciar("Remover duplicados") Then Exit Sub
    On Error GoTo Erro
    Set col = Folhas()
    n = col.Count
    ReDim fx(1 To n): ReDim fy(1 To n): ReDim fw(1 To n): ReDim fh(1 To n)
    ReDim fl(1 To n): ReDim fn(1 To n): ReDim apagar(1 To n): ReDim arr(1 To n)
    i = 0
    For Each s In col
        i = i + 1
        Set arr(i) = s
        s.GetBoundingBox fx(i), fy(i), fw(i), fh(i)
        Set so = s: Set c = Nothing
        On Error Resume Next
        Set c = so.DisplayCurve
        On Error GoTo Erro
        If Not c Is Nothing Then fl(i) = c.Length: fn(i) = c.Nodes.Count
    Next
    For i = 1 To n
        If Not apagar(i) Then
            For j = i + 1 To n
                If Not apagar(j) Then
                    If Abs(fx(i) - fx(j)) <= tol And Abs(fy(i) - fy(j)) <= tol And _
                       Abs(fw(i) - fw(j)) <= tol And Abs(fh(i) - fh(j)) <= tol And _
                       Abs(fl(i) - fl(j)) <= tol * 4 And fn(i) = fn(j) Then apagar(j) = True
                End If
            Next
        End If
    Next
    For i = 1 To n
        If apagar(i) Then arr(i).Delete: del = del + 1
    Next
    Finalizar
    MsgBox del & " objeto(s) duplicado(s) removido(s).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_OtimizarPercurso()
    Dim arr() As Shape, n As Long, i As Long, j As Long, k As Long, mostrar As Double
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double
    Dim feito() As Boolean, ordem() As Long, px As Double, py As Double
    Dim melhor As Long, dmin As Double, d As Double, bloqueado As Boolean
    Dim sx As Double, sy As Double, sw As Double, sh As Double, lay As Layer
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Otimizar percurso", "Corta primeiro os objetos de dentro e depois o mais próximo."
    DlgSim "Mostrar os números da ordem de corte", 1
    If Not DlgMostrar() Then Exit Sub
    mostrar = DV(1)
    If Not Iniciar("Otimizar percurso") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    ReDim bx(1 To n): ReDim by(1 To n): ReDim bw(1 To n): ReDim bh(1 To n)
    ReDim feito(1 To n): ReDim ordem(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox bx(i), by(i), bw(i), bh(i)
    Next
    ActiveSelectionRange.GetBoundingBox sx, sy, sw, sh
    px = sx: py = sy
    For k = 1 To n
        melhor = 0: dmin = 1E+300
        For i = 1 To n
            If Not feito(i) Then
                ' só corta o contorno externo depois dos objetos internos
                bloqueado = False
                For j = 1 To n
                    If j <> i And Not feito(j) Then
                        If CaixaDentro(bx(j), by(j), bw(j), bh(j), bx(i), by(i), bw(i), bh(i)) Then bloqueado = True: Exit For
                    End If
                Next
                If Not bloqueado Then
                    d = (bx(i) + bw(i) / 2 - px) ^ 2 + (by(i) + bh(i) / 2 - py) ^ 2
                    If d < dmin Then dmin = d: melhor = i
                End If
            End If
        Next
        If melhor = 0 Then   ' segurança: pega o primeiro restante
            For i = 1 To n
                If Not feito(i) Then melhor = i: Exit For
            Next
        End If
        feito(melhor) = True: ordem(k) = melhor
        px = bx(melhor) + bw(melhor) / 2: py = by(melhor) + bh(melhor) / 2
    Next
    ' ordem de empilhamento: o 1º a cortar fica atrás, o último na frente
    For k = 1 To n
        arr(ordem(k)).OrderToFront
    Next
    If mostrar = 1 Then
        Set lay = Camada("ORDEM DE CORTE")
        For k = 1 To n
            i = ordem(k)
            Texto(lay, CStr(k), bx(i) + bw(i) / 2, by(i) + bh(i) / 2, 8).Fill.ApplyUniformFill CreateRGBColor(255, 0, 0)
        Next
    End If
    Finalizar
    MsgBox "Ordem de corte otimizada para " & n & " objeto(s)." & vbCr & _
           "Internos primeiro, depois o mais próximo." & vbCr & _
           "O objeto mais ao fundo é o primeiro a ser cortado." & vbCr & _
           IIf(mostrar = 1, "Apague a camada ""ORDEM DE CORTE"" antes de enviar.", ""), vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_FecharCurvas()
    Dim col As Collection, s As Shape, sp As Object, tol As Double, n As Long, abertas As Long, d As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Fechar curvas"
    DlgNum "Fechar se as pontas estiverem a até (mm)", 1
    If Not DlgMostrar() Then Exit Sub
    tol = DV(1)
    If Not Iniciar("Fechar curvas") Then Exit Sub
    On Error GoTo Erro
    Set col = Folhas()
    For Each s In col
        If s.Type = cdrCurveShape Then
            For Each sp In s.Curve.SubPaths
                If Not sp.Closed Then
                    d = Sqr((sp.StartNode.PositionX - sp.EndNode.PositionX) ^ 2 + (sp.StartNode.PositionY - sp.EndNode.PositionY) ^ 2)
                    If d <= tol Then sp.Closed = True: n = n + 1 Else abertas = abertas + 1
                End If
            Next
        End If
    Next
    Finalizar
    MsgBox n & " contorno(s) fechado(s)." & vbCr & abertas & " continuam abertos (pontas mais distantes que a tolerância).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_SimplificarCurvas()
    Dim col As Collection, s As Shape, so As Object, tol As Double, antes As Long, depois As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Simplificar curvas"
    DlgNum "Tolerância (mm) - maior = menos nós", 0.05
    If Not DlgMostrar() Then Exit Sub
    tol = DV(1)
    If Not Iniciar("Simplificar curvas") Then Exit Sub
    On Error GoTo Erro
    Set col = Folhas()
    For Each s In col
        If s.Type = cdrCurveShape Then
            antes = antes + s.Curve.Nodes.Count
            Set so = s
            so.Curve.Nodes.All.AutoReduce tol
            depois = depois + s.Curve.Nodes.Count
        End If
    Next
    Finalizar
    MsgBox "Nós: " & antes & "  ->  " & depois & vbCr & "(somente objetos já convertidos em curva)", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_AgrupamentoAvancado()
    Dim arr() As Shape, n As Long, i As Long, j As Long, raiz() As Long, pai() As Long
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double, melhor As Double
    Dim sr As ShapeRange, grupos As Long
    If Not Iniciar("Agrupamento avançado") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    ReDim bx(1 To n): ReDim by(1 To n): ReDim bw(1 To n): ReDim bh(1 To n)
    ReDim raiz(1 To n): ReDim pai(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox bx(i), by(i), bw(i), bh(i)
    Next
    ' pai = menor objeto que contém a caixa deste
    For i = 1 To n
        melhor = 1E+300
        For j = 1 To n
            If j <> i Then
                If CaixaDentro(bx(i), by(i), bw(i), bh(i), bx(j), by(j), bw(j), bh(j)) Then
                    If bw(j) * bh(j) < melhor Then melhor = bw(j) * bh(j): pai(i) = j
                End If
            End If
        Next
    Next
    For i = 1 To n
        j = i
        Do While pai(j) <> 0
            j = pai(j)
        Loop
        raiz(i) = j
    Next
    For i = 1 To n
        If raiz(i) = i Then
            Set sr = CreateShapeRange
            For j = 1 To n
                If raiz(j) = i Then sr.Add arr(j)
            Next
            If sr.Count > 1 Then sr.Group: grupos = grupos + 1
        End If
    Next
    Finalizar
    MsgBox grupos & " grupo(s) criado(s) (cada peça com os objetos dentro dela).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_SoldarSobrepostos()
    Dim arr() As Shape, n As Long, i As Long, j As Long, g() As Long
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double
    Dim u As Shape, res As Long, mudou As Boolean, a As Long, b As Long, so As Object, sr As ShapeRange
    If Not Iniciar("Soldar sobrepostos") Then Exit Sub
    On Error GoTo Erro
    Set so = ActiveSelectionRange
    Set sr = so.UngroupAllEx
    n = sr.Count
    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = sr(i)
    Next
    ReDim bx(1 To n): ReDim by(1 To n): ReDim bw(1 To n): ReDim bh(1 To n): ReDim g(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox bx(i), by(i), bw(i), bh(i)
        g(i) = i
    Next
    ' agrupa por caixas que se tocam (união transitiva)
    Do
        mudou = False
        For i = 1 To n
            For j = i + 1 To n
                If g(i) <> g(j) Then
                    If bx(i) <= bx(j) + bw(j) And bx(j) <= bx(i) + bw(i) And by(i) <= by(j) + bh(j) And by(j) <= by(i) + bh(i) Then
                        a = IIf(g(i) < g(j), g(i), g(j)): b = IIf(g(i) < g(j), g(j), g(i))
                        For res = 1 To n
                            If g(res) = b Then g(res) = a
                        Next
                        mudou = True
                    End If
                End If
            Next
        Next
    Loop While mudou
    res = 0
    For i = 1 To n
        Set u = Nothing
        For j = 1 To n
            If g(j) = i Then
                If u Is Nothing Then Set u = arr(j) Else Set u = u.Weld(arr(j))
            End If
        Next
        If Not u Is Nothing Then res = res + 1
    Next
    Finalizar
    MsgBox n & " objeto(s) soldados em " & res & " peça(s).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

'==============================================================================
'  3. PLOTTER, LASER E CNC
'==============================================================================
Public Sub LT_ContornoCorte()
    Dim arr() As Shape, n As Long, i As Long, off As Double, base As Shape
    Dim cont As ShapeRange, x As Shape, lay As Layer, tot As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Contorno de corte"
    DlgNum "Distância do contorno (mm, 0 = na arte)", 2
    If Not DlgMostrar() Then Exit Sub
    off = DV(1)
    If Not Iniciar("Contorno de corte") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("CORTE")
    n = Itens(arr)
    For i = 1 To n
        Set base = BaseSolida(arr(i))
        If off > 0 Then
            Set cont = Contornar(base, cdrContourOutside, off)
            base.Delete
        Else
            Set cont = CreateShapeRange: cont.Add base
        End If
        For Each x In cont
            x.Fill.ApplyNoFill
            Estilo x, "corte"
            x.MoveToLayer lay
            tot = tot + 1
        Next
    Next
    Finalizar
    MsgBox tot & " linha(s) de corte criadas em magenta na camada ""CORTE"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_MarcasRegistro()
    Dim dia As Double, marg As Double, qtd As Double, lay As Layer
    Dim x As Double, y As Double, w As Double, h As Double, r As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Marcas de registro"
    DlgNum "Diâmetro da marca (mm)", 5
    DlgNum "Distância até a arte (mm)", 5
    DlgLista "Quantidade de marcas", "3 marcas|4 marcas", 2
    If Not DlgMostrar() Then Exit Sub
    dia = DV(1): marg = DV(2): qtd = DV(3)
    qtd = qtd + 2
    If Not Iniciar("Marcas de registro") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("MARCAS")
    ActiveSelectionRange.GetBoundingBox x, y, w, h
    r = dia / 2
    MarcaPreta lay, x - marg - r, y - marg - r, r
    MarcaPreta lay, x - marg - r, y + h + marg + r, r
    MarcaPreta lay, x + w + marg + r, y + h + marg + r, r
    If qtd >= 4 Then MarcaPreta lay, x + w + marg + r, y - marg - r, r
    Finalizar
    Exit Sub
Erro:
    Falha
End Sub

Private Sub MarcaPreta(lay As Layer, ByVal cx As Double, ByVal cy As Double, ByVal r As Double)
    Dim s As Shape
    Set s = lay.CreateEllipse2(cx, cy, r)
    s.Fill.ApplyUniformFill CreateRGBColor(0, 0, 0)
    s.Outline.SetNoOutline
End Sub

Public Sub LT_LinhasDescarte()
    Dim marg As Double, esp As Double, lay As Layer, s As Shape
    Dim x As Double, y As Double, w As Double, h As Double, p As Double, n As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Linhas de descarte"
    DlgNum "Margem da borda (mm)", 5
    DlgNum "Espaçamento das linhas (mm, 0 = só borda)", 100
    If Not DlgMostrar() Then Exit Sub
    marg = DV(1): esp = DV(2)
    If Not Iniciar("Linhas de descarte") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("CORTE")
    ActiveSelectionRange.GetBoundingBox x, y, w, h
    x = x - marg: y = y - marg: w = w + 2 * marg: h = h + 2 * marg
    Set s = lay.CreateRectangle2(x, y, w, h): Estilo s, "corte"
    If esp > 0 Then
        p = x + esp
        Do While p < x + w - 0.01
            Set s = lay.CreateLineSegment(p, y, p, y + h): Estilo s, "corte"
            p = p + esp: n = n + 1
        Loop
        p = y + esp
        Do While p < y + h - 0.01
            Set s = lay.CreateLineSegment(x, p, x + w, p): Estilo s, "corte"
            p = p + esp: n = n + 1
        Loop
    End If
    Finalizar
    MsgBox "Borda + " & n & " linha(s) de descarte criadas na camada ""CORTE"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_Fatiar()
    Dim tw As Double, th As Double, sobra As Double, sep As Double, rot As Double
    Dim sr As ShapeRange, base As Shape, x As Double, y As Double, w As Double, h As Double
    Dim c As Long, l As Long, nc As Long, nl As Long, r As Shape, p As Shape, tot As Long
    Dim lay As Layer, s As Shape
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Fatiar em painéis"
    DlgNum "Largura de cada painel (mm)", 600
    DlgNum "Altura de cada painel (mm)", 1000
    DlgNum "Sobreposição entre painéis (mm)", 10
    DlgNum "Afastar os painéis (mm)", 20
    DlgSim "Escrever o nome do painel (A1, A2...)", 1
    If Not DlgMostrar() Then Exit Sub
    tw = DV(1): th = DV(2): sobra = DV(3): sep = DV(4)
    rot = DV(5)
    If Not Iniciar("Fatiar") Then Exit Sub
    On Error GoTo Erro
    Set sr = ActiveSelectionRange
    sr.GetBoundingBox x, y, w, h
    Set base = BaseCombinada(sr)
    nc = -Int(-(w - sobra) / (tw - sobra)): If nc < 1 Then nc = 1
    nl = -Int(-(h - sobra) / (th - sobra)): If nl < 1 Then nl = 1
    Set lay = ActiveLayer
    For l = 0 To nl - 1
        For c = 0 To nc - 1
            Set r = lay.CreateRectangle2(x + c * (tw - sobra), y + h - th - l * (th - sobra), tw, th)
            Set p = Nothing
            On Error Resume Next
            Set p = r.Intersect(base, True, True)
            On Error GoTo Erro
            If Not p Is Nothing Then
                p.Move c * sep, -l * sep
                r.Move c * sep, -l * sep
                Estilo r, "marca"
                r.Fill.ApplyNoFill
                If rot = 1 Then
                    Set s = Texto(lay, Chr$(65 + (l Mod 26)) & (c + 1), x + c * (tw - sobra) + c * sep + 15, _
                                  y + h - l * (th - sobra) - l * sep - 10, 14)
                End If
                tot = tot + 1
            Else
                r.Delete
            End If
        Next
    Next
    base.Delete
    Finalizar
    MsgBox tot & " painel(éis) gerados (" & nc & " colunas x " & nl & " linhas)." & vbCr & _
           "O original foi mantido.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_GCode()
    Dim col As Collection, s As Shape, arq As String, fnum As Integer
    Dim prof As Double, passo As Double, zs As Double, fxy As Double, fz As Double, rpm As Double
    Dim sx As Double, sy As Double, sw As Double, sh As Double
    Dim px() As Variant, py() As Variant, fc() As Boolean, dp() As Long, np As Long
    Dim i As Long, j As Long, k As Long, z As Double, ord As Long, nObj As Long
    Dim arr() As Shape, area() As Double, idx() As Long, bw As Double, bh As Double, bx As Double, by As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "G-code para router", "Origem X0 Y0 = canto inferior esquerdo da seleção. Sem compensação de fresa."
    DlgNum "Profundidade total (mm)", 6
    DlgNum "Profundidade por passada (mm)", 2
    DlgNum "Z seguro (mm)", 5
    DlgNum "Avanço XY (mm/min)", 1500
    DlgNum "Mergulho Z (mm/min)", 400
    DlgNum "Rotação do spindle (RPM)", 18000
    DlgTexto "Salvar o arquivo em:", Environ$("USERPROFILE") & "\Desktop\corte_LT.nc"
    If Not DlgMostrar() Then Exit Sub
    prof = DV(1): passo = DV(2): zs = DV(3): fxy = DV(4)
    fz = DV(5): rpm = DV(6): arq = DT(7)
    If passo <= 0 Then passo = prof
    If Not Iniciar("G-code") Then Exit Sub
    On Error GoTo Erro
    ActiveSelectionRange.GetBoundingBox sx, sy, sw, sh
    Set col = Folhas()
    nObj = col.Count
    ReDim arr(1 To nObj): ReDim area(1 To nObj): ReDim idx(1 To nObj)
    i = 0
    For Each s In col
        i = i + 1: Set arr(i) = s: idx(i) = i
        s.GetBoundingBox bx, by, bw, bh
        area(i) = bw * bh            ' menores (internos) primeiro
    Next
    OrdenarIdx idx, area, nObj

    fnum = FreeFile
    Open arq For Output As #fnum
    Print #fnum, "(LT Tools - G-code gerado no CorelDRAW)"
    Print #fnum, "(Origem X0 Y0 = canto inferior esquerdo da seleção)"
    Print #fnum, "(Profundidade " & Fg(prof) & " mm em passadas de " & Fg(passo) & " mm - sem compensação de fresa)"
    Print #fnum, "G21 G90 G17"
    Print #fnum, "G0 Z" & Fg(zs)
    Print #fnum, "M3 S" & CLng(rpm)
    Print #fnum, "G4 P2"
    For k = 1 To nObj
        np = Poligonos(arr(idx(k)), px, py, fc)
        If np > 0 Then
            Profundidades np, px, py, fc, dp
            ' furos (profundidade ímpar) antes do contorno externo
            For ord = 1 To 0 Step -1
                For i = 1 To np
                    If (dp(i) Mod 2) = ord Then
                        Print #fnum, "(contorno)"
                        Print #fnum, "G0 X" & Fg(px(i)(0) - sx) & " Y" & Fg(py(i)(0) - sy)
                        z = 0
                        Do
                            z = z - passo
                            If z < -prof Then z = -prof
                            Print #fnum, "G1 Z" & Fg(z) & " F" & CLng(fz)
                            For j = 1 To UBound(px(i))
                                Print #fnum, "G1 X" & Fg(px(i)(j) - sx) & " Y" & Fg(py(i)(j) - sy) & IIf(j = 1, " F" & CLng(fxy), "")
                            Next
                            If fc(i) Then
                                Print #fnum, "G1 X" & Fg(px(i)(0) - sx) & " Y" & Fg(py(i)(0) - sy)
                            ElseIf z > -prof + 0.0001 Then
                                Print #fnum, "G0 Z" & Fg(zs)
                                Print #fnum, "G0 X" & Fg(px(i)(0) - sx) & " Y" & Fg(py(i)(0) - sy)
                            End If
                        Loop Until z <= -prof + 0.0001
                        Print #fnum, "G0 Z" & Fg(zs)
                    End If
                Next
            Next
        End If
    Next
    Print #fnum, "M5"
    Print #fnum, "G0 X0 Y0"
    Print #fnum, "M30"
    Close #fnum
    Finalizar
    MsgBox "G-code salvo em:" & vbCr & arq & vbCr & vbCr & _
           "Atenção: o caminho segue a linha (sem compensação do diâmetro da fresa)." & vbCr & _
           "Simule o arquivo antes de cortar.", vbInformation, TITULO
    Exit Sub
Erro:
    On Error Resume Next
    Close #fnum
    Falha
End Sub

Public Sub LT_ExportarDXF()
    Dim arq As String, so As Object, ex As Object
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Exportar DXF"
    DlgTexto "Salvar o arquivo DXF em:", Environ$("USERPROFILE") & "\Desktop\corte_LT.dxf"
    If Not DlgMostrar() Then Exit Sub
    arq = DT(1)
    On Error GoTo Erro
    Set so = ActiveDocument
    Set ex = so.ExportEx(arq, cdrDXF, cdrSelection)
    ex.Finish
    MsgBox "DXF salvo em:" & vbCr & arq, vbInformation, TITULO
    Exit Sub
Erro:
    MsgBox "Erro ao exportar: " & Err.Description, vbCritical, TITULO
End Sub

'==============================================================================
'  4. LETRA CAIXA, LED E NEON
'==============================================================================
Public Sub LT_LetraCaixa()
    Dim alt As Double, emenda As Double, maxL As Double, fundo As Double
    Dim arr() As Shape, n As Long, i As Long, c As Object, so As Object, sp As Object
    Dim sx As Double, sy As Double, sw As Double, sh As Double
    Dim cx As Double, cy As Double, L As Double, pedaco As Double, partes As Long, p As Long
    Dim lay As Layer, r As Shape, totL As Double, nf As Long, k As Long, d As Shape
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Letra caixa"
    DlgNum "Profundidade = altura da lateral (mm)", 50
    DlgNum "Sobra para emenda (mm)", 10
    DlgNum "Comprimento máximo da fita (mm, 0 = sem limite)", 1000
    DlgSim "Gerar cópia do fundo", 1
    If Not DlgMostrar() Then Exit Sub
    alt = DV(1): emenda = DV(2): maxL = DV(3): fundo = DV(4)
    If Not Iniciar("Letra caixa") Then Exit Sub
    On Error GoTo Erro
    ActiveSelectionRange.GetBoundingBox sx, sy, sw, sh
    Set lay = Camada("LATERAIS LETRA CAIXA")
    n = Itens(arr)
    cx = sx: cy = sy - 30 - alt
    For i = 1 To n
        Set so = arr(i)
        Set c = so.DisplayCurve
        k = 0
        For Each sp In c.SubPaths
            k = k + 1
            L = sp.Length
            totL = totL + L
            partes = 1
            If maxL > 0 And L + emenda > maxL Then partes = -Int(-L / (maxL - emenda))
            pedaco = L / partes
            For p = 1 To partes
                Set r = lay.CreateRectangle2(cx, cy, pedaco + emenda, alt)
                Estilo r, "corte"
                Texto lay, "Obj " & i & " - contorno " & k & IIf(partes > 1, " (" & p & "/" & partes & ")", "") & _
                           " - " & Fm(pedaco, 1) & " + " & Fm(emenda, 0) & " mm", cx + (pedaco + emenda) / 2, cy + alt / 2, 9
                nf = nf + 1
                cy = cy - alt - 10
            Next
        Next
        If fundo = 1 Then
            Set d = arr(i).Duplicate
            d.Move sw + 30, 0
        End If
    Next
    Finalizar
    MsgBox "Fitas laterais geradas: " & nf & vbCr & _
           "Comprimento total de lateral: " & Fm(totL / 1000#, 2) & " m" & vbCr & _
           "Área de material da lateral: " & Fm((totL + nf * emenda) * alt / 1000000#, 3) & " m²" & _
           IIf(fundo = 1, vbCr & "Cópia do fundo criada à direita.", ""), vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_DistribuirLED()
    Dim mw As Double, mh As Double, ex As Double, ey As Double, marg As Double, pot As Double
    Dim col As Collection, s As Shape, px() As Variant, py() As Variant, fc() As Boolean, np As Long
    Dim x As Double, y As Double, w As Double, h As Double, nx As Long, ny As Long, i As Long, j As Long
    Dim gx As Double, gy As Double, ox As Double, oy As Double, raio As Double
    Dim lay As Layer, led As Shape, tot As Long, fonte As Double, grupo As ShapeRange
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Distribuir LED"
    DlgNum "Largura do módulo (mm)", 15
    DlgNum "Altura do módulo (mm)", 10
    DlgNum "Espaçamento horizontal (mm)", 60
    DlgNum "Espaçamento vertical (mm)", 60
    DlgNum "Distância mínima da borda (mm)", 10
    DlgNum "Potência por módulo (W)", 0.72
    If Not DlgMostrar() Then Exit Sub
    mw = DV(1): mh = DV(2): ex = DV(3): ey = DV(4)
    marg = DV(5): pot = DV(6)
    If Not Iniciar("Distribuir LED") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("LED")
    raio = Sqr(mw * mw + mh * mh) / 2
    Set col = Folhas()
    For Each s In col
        np = Poligonos(s, px, py, fc)
        If np > 0 Then
            s.GetBoundingBox x, y, w, h
            nx = Int((w - 2 * marg) / ex) + 1: If nx < 1 Then nx = 1
            ny = Int((h - 2 * marg) / ey) + 1: If ny < 1 Then ny = 1
            ox = x + (w - (nx - 1) * ex) / 2
            oy = y + (h - (ny - 1) * ey) / 2
            Set grupo = CreateShapeRange
            For j = 0 To ny - 1
                For i = 0 To nx - 1
                    gx = ox + i * ex: gy = oy + j * ey
                    If DentroMaterial(gx, gy, np, px, py, fc) Then
                        If DistBorda(gx, gy, np, px, py) >= marg + raio * 0.5 Then
                            Set led = lay.CreateRectangle2(gx - mw / 2, gy - mh / 2, mw, mh)
                            led.Fill.ApplyUniformFill CorCorel(CorEstilo(E_LED, 7))
                            led.Outline.Width = 0.1
                            grupo.Add led
                            tot = tot + 1
                        End If
                    End If
                Next
            Next
            If grupo.Count > 1 Then grupo.Group
        End If
    Next
    Finalizar
    fonte = FonteComercial(tot * pot * 1.2)
    MsgBox "Módulos de LED: " & tot & vbCr & _
           "Potência total: " & Fm(tot * pot, 1) & " W" & vbCr & _
           "Fonte sugerida (+20% folga): " & fonte & " W" & vbCr & vbCr & _
           "Confira visualmente - em traços finos pode ser preciso ajustar à mão.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_Neon()
    Dim dia As Double, corN As Double, pm As Double, wm As Double
    Dim col As Collection, s As Shape, so As Object, c As Object, L As Double, segs As Long
    Dim r As Long, g As Long, b As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Neon"
    DlgNum "Largura do neon (mm)", 8
    DlgLista "Cor", "Branco frio|Branco quente|Vermelho|Verde|Azul|Rosa|Amarelo|Laranja|Roxo|Gelo", 6
    DlgNum "Preço (R$ por metro)", 25
    DlgNum "Consumo (W por metro)", 10
    If Not DlgMostrar() Then Exit Sub
    dia = DV(1): corN = DV(2): pm = DV(3): wm = DV(4)
    Select Case corN
        Case 1: r = 230: g = 240: b = 255
        Case 2: r = 255: g = 214: b = 160
        Case 3: r = 255: g = 30: b = 30
        Case 4: r = 30: g = 230: b = 60
        Case 5: r = 30: g = 90: b = 255
        Case 6: r = 255: g = 60: b = 200
        Case 7: r = 255: g = 240: b = 40
        Case 8: r = 255: g = 140: b = 20
        Case 9: r = 160: g = 50: b = 255
        Case Else: r = 180: g = 240: b = 255
    End Select
    If Not Iniciar("Neon") Then Exit Sub
    On Error GoTo Erro
    Set col = Folhas()
    For Each s In col
        Set so = s
        Set c = Nothing
        On Error Resume Next
        Set c = so.DisplayCurve
        On Error GoTo Erro
        If Not c Is Nothing Then
            L = L + c.Length
            segs = segs + c.SubPaths.Count
            s.Fill.ApplyNoFill
            s.Outline.Width = dia
            s.Outline.Color.RGBAssign r, g, b
            s.Outline.LineCaps = cdrOutlineRoundLineCaps
            s.Outline.LineJoin = cdrOutlineRoundLineJoin
        End If
    Next
    Finalizar
    MsgBox "Metragem de neon: " & Fm(L / 1000#, 2) & " m" & vbCr & _
           "Trechos (pontas / emendas): " & segs & vbCr & _
           "Custo do neon: R$ " & Fm(L / 1000# * pm) & vbCr & _
           "Consumo: " & Fm(L / 1000# * wm, 1) & " W  ->  fonte sugerida " & FonteComercial(L / 1000# * wm * 1.2) & " W", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_FurosContorno()
    Dim esp As Double, dia As Double, recuo As Double, modo As Double
    Dim col As Collection, s As Shape, px() As Variant, py() As Variant, fc() As Boolean, dp() As Long, np As Long
    Dim i As Long, j As Long, m As Long, ori As Double, perim As Double, alvo As Double, acum As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, L As Double, t As Double, qtd As Long
    Dim tx As Double, ty As Double, nx As Double, ny As Double, cx As Double, cy As Double, passo As Double
    Dim lay As Layer, f As Shape, tot As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Furos no contorno"
    DlgNum "Espaçamento entre furos (mm)", 150
    DlgNum "Diâmetro do furo (mm)", 4
    DlgNum "Recuo para dentro da peça (mm)", 8
    DlgLista "Onde furar", "Contorno inteiro (montagem)|Só na parte de baixo (drenagem)", 1
    If Not DlgMostrar() Then Exit Sub
    esp = DV(1): dia = DV(2): recuo = DV(3): modo = DV(4)
    If Not Iniciar("Furos no contorno") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("FUROS")
    Set col = Folhas()
    For Each s In col
        np = Poligonos(s, px, py, fc)
        If np > 0 Then
            Profundidades np, px, py, fc, dp
            For i = 1 To np
                If fc(i) And dp(i) Mod 2 = 0 Then
                    ori = IIf(AreaSinal(px(i), py(i)) >= 0, 1, -1)
                    m = UBound(px(i)) + 1
                    perim = 0
                    For j = 0 To m - 1
                        perim = perim + Sqr((px(i)((j + 1) Mod m) - px(i)(j)) ^ 2 + (py(i)((j + 1) Mod m) - py(i)(j)) ^ 2)
                    Next
                    qtd = -Int(-perim / esp): If qtd < 1 Then qtd = 1
                    passo = perim / qtd
                    alvo = passo / 2: acum = 0
                    For j = 0 To m - 1
                        x1 = px(i)(j): y1 = py(i)(j)
                        x2 = px(i)((j + 1) Mod m): y2 = py(i)((j + 1) Mod m)
                        L = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
                        Do While L > 0 And alvo <= acum + L
                            t = (alvo - acum) / L
                            tx = (x2 - x1) / L: ty = (y2 - y1) / L
                            nx = -ty * ori: ny = tx * ori      ' normal para dentro
                            cx = x1 + t * (x2 - x1) + nx * recuo
                            cy = y1 + t * (y2 - y1) + ny * recuo
                            ' modo drenagem: só onde a borda está virada para baixo
                            If modo <> 2 Or ny > 0.7 Then
                                Set f = lay.CreateEllipse2(cx, cy, dia / 2)
                                Estilo f, "furo"
                                tot = tot + 1
                            End If
                            alvo = alvo + passo
                        Loop
                        acum = acum + L
                    Next
                End If
            Next
        End If
    Next
    Finalizar
    MsgBox tot & " furo(s) de Ø" & Fm(dia, 1) & " mm criados na camada ""FUROS"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_Ilhoses()
    Dim esp As Double, borda As Double, dia As Double, arr() As Shape, n As Long, i As Long
    Dim x As Double, y As Double, w As Double, h As Double, nx As Long, ny As Long, k As Long
    Dim lay As Layer, tot As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Ilhoses"
    DlgNum "Espaçamento máximo (mm)", 300
    DlgNum "Distância da borda (mm)", 15
    DlgNum "Diâmetro da marca (mm)", 10
    If Not DlgMostrar() Then Exit Sub
    esp = DV(1): borda = DV(2): dia = DV(3)
    If Not Iniciar("Ilhoses") Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("ILHOSES")
    n = Itens(arr)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        x = x + borda: y = y + borda: w = w - 2 * borda: h = h - 2 * borda
        nx = -Int(-w / esp): If nx < 1 Then nx = 1
        ny = -Int(-h / esp): If ny < 1 Then ny = 1
        For k = 0 To nx
            Ilhos lay, x + k * w / nx, y, dia: Ilhos lay, x + k * w / nx, y + h, dia
            tot = tot + 2
        Next
        For k = 1 To ny - 1
            Ilhos lay, x, y + k * h / ny, dia: Ilhos lay, x + w, y + k * h / ny, dia
            tot = tot + 2
        Next
    Next
    Finalizar
    MsgBox tot & " ilhós(es) marcados.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Private Sub Ilhos(lay As Layer, ByVal cx As Double, ByVal cy As Double, ByVal dia As Double)
    Dim s As Shape
    Set s = lay.CreateEllipse2(cx, cy, dia / 2)
    s.Fill.ApplyNoFill
    Estilo s, "furo"
    Set s = lay.CreateEllipse2(cx, cy, dia / 4)
    Estilo s, "furo"
End Sub

'==============================================================================
'  5. GERADORES
'==============================================================================
Public Sub LT_CaixaMalhete()
    Dim X As Double, Y As Double, Z As Double, t As Double, f As Double, tampa As Double
    Dim ox As Double, oy As Double, gap As Double, lay As Layer, eT As Integer
    DlgNovo "Caixa com malhete", "Medidas externas. A caixa é desenhada a partir do ponto 0,0 da página."
    DlgNum "Comprimento (mm)", 200
    DlgNum "Largura (mm)", 150
    DlgNum "Altura (mm)", 100
    DlgNum "Espessura do material (mm)", 3
    DlgNum "Largura do dente (mm)", 15
    DlgSim "Com tampa", 0
    If Not DlgMostrar() Then Exit Sub
    X = DV(1): Y = DV(2): Z = DV(3): t = DV(4)
    f = DV(5): tampa = DV(6)
    If t <= 0 Or f <= t Then MsgBox "Dente precisa ser maior que a espessura.", vbExclamation, TITULO: Exit Sub
    If Not Iniciar("Caixa com malhete", False) Then Exit Sub
    On Error GoTo Erro
    Set lay = ActiveLayer
    gap = 10
    eT = IIf(tampa = 1, 2, 0)
    ox = 0: oy = 0
    ' tipos de borda: 0 lisa, 1 dente para fora (macho), 2 dente para dentro (fêmea)
    Painel lay, "FUNDO", X, Y, 1, 1, 1, 1, t, f, ox, oy: ox = ox + X + gap
    If tampa = 1 Then Painel lay, "TAMPA", X, Y, 1, 1, 1, 1, t, f, ox, oy: ox = ox + X + gap
    Painel lay, "FRENTE", X, Z, 2, 1, eT, 1, t, f, ox, oy: ox = ox + X + gap
    Painel lay, "TRÁS", X, Z, 2, 1, eT, 1, t, f, ox, oy: ox = ox + X + gap
    Painel lay, "LATERAL", Y, Z, 2, 2, eT, 2, t, f, ox, oy: ox = ox + Y + gap
    Painel lay, "LATERAL", Y, Z, 2, 2, eT, 2, t, f, ox, oy
    Finalizar
    MsgBox "Caixa gerada a partir do ponto 0,0 da página." & vbCr & _
           "Para laser, ajuste o kerf (compensação do feixe) no software da máquina.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Private Sub Painel(lay As Layer, ByVal nome As String, ByVal W As Double, ByVal H As Double, _
                   ByVal eB As Integer, ByVal eR As Integer, ByVal eT As Integer, ByVal eL As Integer, _
                   ByVal t As Double, ByVal f As Double, ByVal ox As Double, ByVal oy As Double)
    Dim xs() As Double, ys() As Double, m As Long, k As Long
    Dim nW As Long, nH As Long, wW As Double, wH As Double, a As Double, b As Double
    Dim dB As Double, dR As Double, dT As Double, dL As Double, s As Shape
    ReDim xs(0 To 63): ReDim ys(0 To 63)
    nW = NSeg(W, f): nH = NSeg(H, f)
    wW = W / nW: wH = H / nH
    dB = Dep(eB, 0, t): dR = Dep(eR, 0, t): dT = Dep(eT, 0, t): dL = Dep(eL, 0, t)
    AddPt xs, ys, m, dL, dB
    For k = 0 To nW - 2                                   ' base  ->
        a = Dep(eB, k, t): b = Dep(eB, k + 1, t)
        If a <> b Then AddPt xs, ys, m, (k + 1) * wW, a: AddPt xs, ys, m, (k + 1) * wW, b
    Next
    AddPt xs, ys, m, W - dR, dB
    For k = 0 To nH - 2                                   ' direita ^
        a = Dep(eR, k, t): b = Dep(eR, k + 1, t)
        If a <> b Then AddPt xs, ys, m, W - a, (k + 1) * wH: AddPt xs, ys, m, W - b, (k + 1) * wH
    Next
    AddPt xs, ys, m, W - dR, H - dT
    For k = 0 To nW - 2                                   ' topo <-
        a = Dep(eT, k, t): b = Dep(eT, k + 1, t)
        If a <> b Then AddPt xs, ys, m, W - (k + 1) * wW, H - a: AddPt xs, ys, m, W - (k + 1) * wW, H - b
    Next
    AddPt xs, ys, m, dL, H - dT
    For k = 0 To nH - 2                                   ' esquerda v
        a = Dep(eL, k, t): b = Dep(eL, k + 1, t)
        If a <> b Then AddPt xs, ys, m, a, H - (k + 1) * wH: AddPt xs, ys, m, b, H - (k + 1) * wH
    Next
    Set s = CriarPoligono(lay, xs, ys, m, ox, oy)
    Estilo s, "corte"
    s.Name = nome
    Texto lay, nome, ox + W / 2, oy + H / 2, 10
End Sub

Private Function NSeg(ByVal L As Double, ByVal f As Double) As Long
    Dim n As Long
    n = Int(L / f)
    If n Mod 2 = 0 Then n = n - 1
    If n < 3 Then n = 3
    NSeg = n
End Function

Private Function Dep(ByVal tipo As Integer, ByVal k As Long, ByVal t As Double) As Double
    Select Case tipo
        Case 1: If k Mod 2 = 0 Then Dep = 0 Else Dep = t
        Case 2: If k Mod 2 = 0 Then Dep = t Else Dep = 0
        Case Else: Dep = 0
    End Select
End Function

Public Sub LT_QuebraCabeca()
    Dim cols As Double, lins As Double, x As Double, y As Double, w As Double, h As Double
    Dim cw As Double, ch As Double, T As Double, c As Long, r As Long, lay As Layer
    Dim crv As Object, sp As Object, s As Shape
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    DlgNovo "Quebra-cabeça", "Usa o tamanho do objeto selecionado. Sem seleção, usa a largura e altura abaixo."
    DlgNum "Colunas de peças", 5
    DlgNum "Linhas de peças", 4
    DlgNum "Largura (sem seleção, mm)", 300
    DlgNum "Altura (sem seleção, mm)", 200
    If Not DlgMostrar() Then Exit Sub
    cols = DV(1): lins = DV(2)
    If cols < 1 Or lins < 1 Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then
        w = DV(3): h = DV(4)
        x = 0: y = 0
        If Not Iniciar("Quebra-cabeça", False) Then Exit Sub
    Else
        If Not Iniciar("Quebra-cabeça") Then Exit Sub
        ActiveSelectionRange.GetBoundingBox x, y, w, h
    End If
    On Error GoTo Erro
    Randomize
    Set lay = Camada("CORTE")
    cw = w / cols: ch = h / lins
    T = IIf(cw < ch, cw, ch)
    ' linhas verticais
    For c = 1 To cols - 1
        Set crv = ActiveDocument.CreateCurve
        Set sp = crv.CreateSubPath(x + c * cw, y)
        For r = 0 To lins - 1
            Aba sp, x + c * cw, y + r * ch, 0, 1, ch, T, IIf(Rnd < 0.5, 1, -1)
        Next
        Set s = lay.CreateCurve(crv): Estilo s, "corte"
    Next
    ' linhas horizontais
    For r = 1 To lins - 1
        Set crv = ActiveDocument.CreateCurve
        Set sp = crv.CreateSubPath(x, y + r * ch)
        For c = 0 To cols - 1
            Aba sp, x + c * cw, y + r * ch, 1, 0, cw, T, IIf(Rnd < 0.5, 1, -1)
        Next
        Set s = lay.CreateCurve(crv): Estilo s, "corte"
    Next
    Set s = lay.CreateRectangle2(x, y, w, h): Estilo s, "corte"
    Finalizar
    MsgBox cols * lins & " peças geradas na camada ""CORTE"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

' acrescenta uma borda com encaixe (aba) a partir de (ax,ay) na direção (ux,uy)
Private Sub Aba(sp As Object, ByVal ax As Double, ByVal ay As Double, ByVal ux As Double, ByVal uy As Double, _
                ByVal L As Double, ByVal T As Double, ByVal sinal As Double)
    gAx = ax: gAy = ay: gUx = ux: gUy = uy: gL = L: gT = T
    gNx = -uy * sinal: gNy = ux * sinal
    sp.AppendLineSegment PtX(0.38, 0), PtY(0.38, 0)
    sp.AppendCurveSegment2 PtX(0.42, 0.1), PtY(0.42, 0.1), PtX(0.44, 0), PtY(0.44, 0), PtX(0.44, 0.06), PtY(0.44, 0.06)
    sp.AppendCurveSegment2 PtX(0.5, 0.32), PtY(0.5, 0.32), PtX(0.36, 0.18), PtY(0.36, 0.18), PtX(0.4, 0.32), PtY(0.4, 0.32)
    sp.AppendCurveSegment2 PtX(0.58, 0.1), PtY(0.58, 0.1), PtX(0.6, 0.32), PtY(0.6, 0.32), PtX(0.64, 0.18), PtY(0.64, 0.18)
    sp.AppendCurveSegment2 PtX(0.62, 0), PtY(0.62, 0), PtX(0.56, 0.06), PtY(0.56, 0.06), PtX(0.56, 0), PtY(0.56, 0)
    sp.AppendLineSegment ax + ux * L, ay + uy * L
End Sub

Private Function PtX(ByVal a As Double, ByVal hh As Double) As Double
    PtX = gAx + gUx * (gL / 2 + (a - 0.5) * gT) + gNx * hh * gT
End Function

Private Function PtY(ByVal a As Double, ByVal hh As Double) As Double
    PtY = gAy + gUy * (gL / 2 + (a - 0.5) * gT) + gNy * hh * gT
End Function

Public Sub LT_Grelha()
    Dim W As Double, H As Double, cols As Double, lins As Double, gap As Double
    Dim x As Double, y As Double, c As Long, l As Long, cw As Double, ch As Double
    Dim lay As Layer, s As Shape, sr As ShapeRange
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    DlgNovo "Grelha", "Usa o tamanho do objeto selecionado. Sem seleção, usa a largura e altura abaixo."
    DlgNum "Colunas", 4
    DlgNum "Linhas", 2
    DlgNum "Espaço entre peças (mm, 0 = encostadas)", 0
    DlgNum "Largura total (sem seleção, mm)", 1000
    DlgNum "Altura total (sem seleção, mm)", 500
    If Not DlgMostrar() Then Exit Sub
    cols = DV(1): lins = DV(2): gap = DV(3)
    If ActiveSelectionRange.Count > 0 Then
        ActiveSelectionRange.GetBoundingBox x, y, W, H
    Else
        W = DV(4): H = DV(5)
    End If
    If cols < 1 Or lins < 1 Then Exit Sub
    If Not Iniciar("Grelha", False) Then Exit Sub
    On Error GoTo Erro
    Set lay = Camada("CORTE")
    cw = (W - (cols - 1) * gap) / cols: ch = (H - (lins - 1) * gap) / lins
    Set sr = CreateShapeRange
    For l = 0 To lins - 1
        For c = 0 To cols - 1
            Set s = lay.CreateRectangle2(x + c * (cw + gap), y + l * (ch + gap), cw, ch)
            Estilo s, "corte"
            sr.Add s
        Next
    Next
    sr.Group
    Finalizar
    MsgBox cols * lins & " peças de " & Fm(cw, 1) & " x " & Fm(ch, 1) & " mm.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_CirculosAleatorios()
    Dim rmin As Double, rmax As Double, gap As Double, tent As Double
    Dim col As Collection, s As Shape, px() As Variant, py() As Variant, fc() As Boolean, np As Long
    Dim x As Double, y As Double, w As Double, h As Double, i As Long, k As Long
    Dim cx() As Double, cy() As Double, cr() As Double, nc As Long, gx As Double, gy As Double
    Dim disp As Double, d As Double, lay As Layer, e As Shape, sr As ShapeRange, tot As Long
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Círculos aleatórios"
    DlgNum "Raio mínimo (mm)", 2
    DlgNum "Raio máximo (mm)", 12
    DlgNum "Espaço entre círculos (mm)", 1.5
    DlgNum "Tentativas (mais = mais cheio)", 3000
    If Not DlgMostrar() Then Exit Sub
    rmin = DV(1): rmax = DV(2): gap = DV(3): tent = DV(4)
    If Not Iniciar("Círculos aleatórios") Then Exit Sub
    On Error GoTo Erro
    Randomize
    Set lay = ActiveLayer
    Set col = Folhas()
    For Each s In col
        np = Poligonos(s, px, py, fc)
        If np > 0 Then
            s.GetBoundingBox x, y, w, h
            nc = 0: ReDim cx(1 To 1000): ReDim cy(1 To 1000): ReDim cr(1 To 1000)
            Set sr = CreateShapeRange
            For k = 1 To tent
                gx = x + Rnd * w: gy = y + Rnd * h
                If DentroMaterial(gx, gy, np, px, py, fc) Then
                    disp = DistBorda(gx, gy, np, px, py) - gap
                    For i = 1 To nc
                        d = Sqr((gx - cx(i)) ^ 2 + (gy - cy(i)) ^ 2) - cr(i) - gap
                        If d < disp Then disp = d
                        If disp < rmin Then Exit For
                    Next
                    If disp >= rmin Then
                        If disp > rmax Then disp = rmax
                        nc = nc + 1
                        If nc > UBound(cx) Then ReDim Preserve cx(1 To nc * 2): ReDim Preserve cy(1 To nc * 2): ReDim Preserve cr(1 To nc * 2)
                        cx(nc) = gx: cy(nc) = gy: cr(nc) = disp
                        Set e = lay.CreateEllipse2(gx, gy, disp)
                        Estilo e, "corte"
                        sr.Add e
                    End If
                End If
            Next
            If sr.Count > 1 Then sr.Group
            tot = tot + nc
        End If
    Next
    Finalizar
    MsgBox tot & " círculo(s) criados.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_NumerarObjetos()
    Dim pre As String, ini As Double, dig As Double, tam As Double, ordem As Double
    Dim arr() As Shape, n As Long, i As Long, j As Long, idx() As Long, v As Long
    Dim cx() As Double, cy() As Double, hh() As Double, x As Double, y As Double, w As Double, h As Double
    Dim tol As Double, lay As Layer, txt As String
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Numerar objetos"
    DlgTexto "Prefixo (ex.: PEÇA )", ""
    DlgNum "Número inicial", 1
    DlgNum "Quantidade de dígitos (3 = 001)", 1
    DlgNum "Tamanho do texto (pt)", 12
    DlgLista "Ordem", "Linhas (esquerda para direita)|Colunas (cima para baixo)", 1
    If Not DlgMostrar() Then Exit Sub
    pre = DT(1): ini = DV(2): dig = DV(3): tam = DV(4)
    ordem = DV(5)
    If Not Iniciar("Numerar objetos") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    ReDim cx(1 To n): ReDim cy(1 To n): ReDim hh(1 To n): ReDim idx(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        cx(i) = x + w / 2: cy(i) = y + h / 2: hh(i) = IIf(ordem = 1, h, w): idx(i) = i
        tol = tol + hh(i)
    Next
    tol = tol / n / 2
    ' ordenação por inserção com tolerância de linha/coluna
    For i = 2 To n
        v = idx(i): j = i - 1
        Do While j >= 1
            If Not VemAntes(v, idx(j), cx, cy, tol, ordem) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = v
    Next
    Set lay = Camada("NUMERAÇÃO")
    For i = 1 To n
        txt = pre & Format$(ini + i - 1, String$(IIf(dig < 1, 1, dig), "0"))
        Texto lay, txt, cx(idx(i)), cy(idx(i)), tam
    Next
    Finalizar
    MsgBox n & " objeto(s) numerados na camada ""NUMERAÇÃO"".", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Private Function VemAntes(ByVal a As Long, ByVal b As Long, cx() As Double, cy() As Double, ByVal tol As Double, ByVal ordem As Double) As Boolean
    If ordem = 1 Then
        If Abs(cy(a) - cy(b)) > tol Then VemAntes = cy(a) > cy(b) Else VemAntes = cx(a) < cx(b)
    Else
        If Abs(cx(a) - cx(b)) > tol Then VemAntes = cx(a) < cx(b) Else VemAntes = cy(a) > cy(b)
    End If
End Function

Public Sub LT_SequenciaNumerada()
    Dim ini As Double, qtd As Double, cols As Double, gap As Double
    Dim sr As ShapeRange, d As ShapeRange, x As Double, y As Double, w As Double, h As Double
    Dim i As Long, c As Long, l As Long, col As Collection, s As Shape, so As Object
    Dim txt As String, p As Long, k As Long, achou As Boolean
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If Not gConfig Then
        If ActiveSelectionRange.Count = 0 Then
            MsgBox "Selecione o modelo (arte + texto com ### no lugar do número)." & vbCr & _
                   "Ex.: ""Nº ###"" vira Nº 001, Nº 002...", vbExclamation, TITULO
            Exit Sub
        End If
    End If
    DlgNovo "Sequência numerada", "O texto do modelo deve ter ### no lugar do número."
    DlgNum "Primeiro número", 1
    DlgNum "Quantidade", 50
    DlgNum "Colunas", 5
    DlgNum "Espaço entre cópias (mm)", 2
    If Not DlgMostrar() Then Exit Sub
    ini = DV(1): qtd = DV(2): cols = DV(3): gap = DV(4)
    If cols < 1 Then cols = 1
    If Not Iniciar("Sequência numerada") Then Exit Sub
    On Error GoTo Erro
    Set sr = ActiveSelectionRange
    sr.GetBoundingBox x, y, w, h
    For i = 0 To qtd - 1
        c = i Mod cols: l = i \ cols
        Set d = sr.Duplicate
        d.Move (c + 1) * (w + gap), -l * (h + gap)
        Set col = New Collection
        ColetarFolhas d, col
        For Each s In col
            If s.Type = cdrTextShape Then
                Set so = s
                txt = so.Text.Story.Text
                p = InStr(txt, "#")
                If p > 0 Then
                    k = 0
                    Do While p + k <= Len(txt)
                        If Mid$(txt, p + k, 1) <> "#" Then Exit Do
                        k = k + 1
                    Loop
                    so.Text.Story.Text = Left$(txt, p - 1) & Format$(ini + i, String$(k, "0")) & Mid$(txt, p + k)
                    achou = True
                End If
            End If
        Next
    Next
    Finalizar
    If Not achou Then MsgBox "Nenhum texto com ### foi encontrado no modelo - as cópias foram feitas sem número.", vbExclamation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_CaixaDelimitadora()
    Dim marg As Double, modo As Double, raio As Double, arr() As Shape, n As Long, i As Long
    Dim x As Double, y As Double, w As Double, h As Double, s As Shape
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Caixa delimitadora"
    DlgNum "Margem em volta (mm)", 5
    DlgLista "Criar", "Uma caixa para cada objeto|Uma caixa para a seleção toda", 1
    DlgNum "Raio dos cantos (mm)", 0
    If Not DlgMostrar() Then Exit Sub
    marg = DV(1): modo = DV(2): raio = DV(3)
    If Not Iniciar("Caixa delimitadora") Then Exit Sub
    On Error GoTo Erro
    If modo = 2 Then
        ActiveSelectionRange.GetBoundingBox x, y, w, h
        CaixaEm x - marg, y - marg, w + 2 * marg, h + 2 * marg, raio
    Else
        n = Itens(arr)
        For i = 1 To n
            arr(i).GetBoundingBox x, y, w, h
            CaixaEm x - marg, y - marg, w + 2 * marg, h + 2 * marg, raio
        Next
    End If
    Finalizar
    Exit Sub
Erro:
    Falha
End Sub

Private Sub CaixaEm(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal raio As Double)
    Dim s As Shape
    Set s = ActiveLayer.CreateRectangle2(x, y, w, h, raio, raio, raio, raio)
    Estilo s, "corte"
End Sub

'==============================================================================
'  6. EDIÇÃO DE OBJETOS
'==============================================================================
Public Sub LT_DefinirTamanho()
    Dim nw As Double, nh As Double, arr() As Shape, n As Long, i As Long
    Dim x As Double, y As Double, w As Double, h As Double, fw As Double, fh As Double, x2 As Double, y2 As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Definir tamanho", "Deixe 0 em um dos campos para manter a proporção."
    DlgNum "Nova largura (mm)", 0
    DlgNum "Nova altura (mm)", 0
    If Not DlgMostrar() Then Exit Sub
    nw = DV(1): nh = DV(2)
    If nw <= 0 And nh <= 0 Then Exit Sub
    If Not Iniciar("Definir tamanho") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        If w > 0 And h > 0 Then
            fw = nw: fh = nh
            If fw <= 0 Then fw = w * nh / h
            If fh <= 0 Then fh = h * nw / w
            arr(i).SetSize fw, fh
            arr(i).GetBoundingBox x2, y2, w, h
            arr(i).Move x - x2, y - y2          ' mantém o canto inferior esquerdo
        End If
    Next
    Finalizar
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_DistribuirEspaco()
    Dim dirc As Double, gap As Double, arr() As Shape, n As Long, i As Long, idx() As Long, chave() As Double
    Dim x As Double, y As Double, w As Double, h As Double, pos As Double
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Distribuir com espaçamento"
    DlgLista "Direção", "Horizontal (lado a lado)|Vertical (um embaixo do outro)", 1
    DlgNum "Espaço entre objetos (mm)", 5
    If Not DlgMostrar() Then Exit Sub
    dirc = DV(1): gap = DV(2)
    If Not Iniciar("Distribuir com espaço") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    If n < 2 Then Finalizar: Exit Sub
    ReDim idx(1 To n): ReDim chave(1 To n)
    For i = 1 To n
        arr(i).GetBoundingBox x, y, w, h
        idx(i) = i
        chave(i) = IIf(dirc = 1, x, -(y + h))
    Next
    OrdenarIdx idx, chave, n
    arr(idx(1)).GetBoundingBox x, y, w, h
    If dirc = 1 Then pos = x + w + gap Else pos = y - gap
    For i = 2 To n
        arr(idx(i)).GetBoundingBox x, y, w, h
        If dirc = 1 Then
            arr(idx(i)).Move pos - x, 0
            pos = pos + w + gap
        Else
            arr(idx(i)).Move 0, pos - (y + h)
            pos = pos - h - gap
        End If
    Next
    Finalizar
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_ObjetosAoLongoCurva()
    Dim arr() As Shape, n As Long, caminho As Shape, modelo As Shape
    Dim esp As Double, qtd As Double, girar As Double, x As Double, y As Double, w As Double, h As Double
    Dim a1 As Double, a2 As Double, px() As Variant, py() As Variant, fc() As Boolean, np As Long
    Dim m As Long, j As Long, perim As Double, passo As Double, alvo As Double, acum As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, L As Double, t As Double
    Dim cx As Double, cy As Double, d As Shape, tot As Long, ultimo As Long
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If Not gConfig Then
        If ActiveSelectionRange.Count <> 2 Then
            MsgBox "Selecione 2 objetos: o caminho (o maior) e o objeto que será repetido.", vbExclamation, TITULO
            Exit Sub
        End If
    End If
    DlgNovo "Objetos ao longo da curva", "Selecione 2 objetos: o caminho (o maior) e o objeto a repetir."
    DlgLista "Distribuir por", "Quantidade de cópias|Espaçamento fixo", 1
    DlgNum "Quantidade de cópias", 10
    DlgNum "Espaçamento (mm)", 50
    DlgSim "Girar acompanhando a curva", 1
    If Not DlgMostrar() Then Exit Sub
    qtd = DV(2): esp = DV(3): girar = DV(4)
    If DV(1) = 1 Then esp = 0
    If Not Iniciar("Objetos ao longo da curva") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    arr(1).GetBoundingBox x, y, w, h: a1 = w * h
    arr(2).GetBoundingBox x, y, w, h: a2 = w * h
    If a1 >= a2 Then Set caminho = arr(1): Set modelo = arr(2) Else Set caminho = arr(2): Set modelo = arr(1)
    np = Poligonos(caminho, px, py, fc)
    If np = 0 Then Err.Raise 1000, , "O caminho não é uma curva."
    m = UBound(px(1)) + 1
    ultimo = IIf(fc(1), m, m - 1)
    For j = 0 To ultimo - 1
        perim = perim + Sqr((px(1)((j + 1) Mod m) - px(1)(j)) ^ 2 + (py(1)((j + 1) Mod m) - py(1)(j)) ^ 2)
    Next
    If esp > 0 Then
        passo = esp
    Else
        If qtd < 1 Then qtd = 1
        passo = IIf(fc(1), perim / qtd, perim / IIf(qtd > 1, qtd - 1, 1))
    End If
    alvo = 0: acum = 0
    For j = 0 To ultimo - 1
        x1 = px(1)(j): y1 = py(1)(j)
        x2 = px(1)((j + 1) Mod m): y2 = py(1)((j + 1) Mod m)
        L = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
        Do While L > 0 And alvo <= acum + L + 0.0001
            If fc(1) And alvo >= perim - 0.0001 Then Exit Do
            t = (alvo - acum) / L: If t > 1 Then t = 1
            cx = x1 + t * (x2 - x1): cy = y1 + t * (y2 - y1)
            Set d = modelo.Duplicate
            d.GetBoundingBox x, y, w, h
            d.Move cx - (x + w / 2), cy - (y + h / 2)
            If girar = 1 Then d.RotateEx Atan2(y2 - y1, x2 - x1) * 180 / PI, cx, cy
            tot = tot + 1
            alvo = alvo + passo
        Loop
        acum = acum + L
    Next
    Finalizar
    MsgBox tot & " cópia(s) distribuídas ao longo da curva.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_TrocarObjetos()
    Dim modelo As Shape, arr() As Shape, n As Long, i As Long, ajustar As Double
    Dim x As Double, y As Double, w As Double, h As Double, mx As Double, my As Double, mw As Double, mh As Double
    Dim d As Shape, f As Double, tot As Long
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If Not gConfig Then
        If ActiveSelectionRange.Count < 2 Then
            MsgBox "Selecione os objetos que serão trocados e, POR ÚLTIMO, o modelo (Shift+clique).", vbExclamation, TITULO
            Exit Sub
        End If
    End If
    DlgNovo "Trocar objetos", "Selecione os objetos e, por último, o modelo."
    DlgSim "Ajustar o modelo ao tamanho de cada objeto", 0
    If Not DlgMostrar() Then Exit Sub
    ajustar = DV(1)
    If Not Iniciar("Trocar objetos") Then Exit Sub
    On Error GoTo Erro
    Set modelo = ActiveShape
    modelo.GetBoundingBox mx, my, mw, mh
    n = Itens(arr)
    For i = 1 To n
        If arr(i).StaticID <> modelo.StaticID Then
            arr(i).GetBoundingBox x, y, w, h
            Set d = modelo.Duplicate
            If ajustar = 1 And mw > 0 And mh > 0 Then
                f = IIf(w / mw < h / mh, w / mw, h / mh)
                d.SetSize mw * f, mh * f
            End If
            d.GetBoundingBox mx, my, mw, mh
            d.Move (x + w / 2) - (mx + mw / 2), (y + h / 2) - (my + mh / 2)
            arr(i).Delete
            tot = tot + 1
            modelo.GetBoundingBox mx, my, mw, mh
        End If
    Next
    Finalizar
    MsgBox tot & " objeto(s) trocados pelo modelo.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_EspelharPorLinha()
    Dim arr() As Shape, n As Long, i As Long, linha As Shape, copia As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, ang As Double
    Dim s As Shape, x As Double, y As Double, w As Double, h As Double, tot As Long
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If Not gConfig Then
        If ActiveSelectionRange.Count < 2 Then
            MsgBox "Selecione os objetos e uma linha reta (2 nós) que será o eixo do espelho.", vbExclamation, TITULO
            Exit Sub
        End If
    End If
    DlgNovo "Espelhar por linha"
    DlgSim "Manter o original (espelhar uma cópia)", 1
    If Not DlgMostrar() Then Exit Sub
    copia = DV(1)
    If Not Iniciar("Espelhar por linha") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    For i = 1 To n
        If EhLinha(arr(i), x0, y0, x1, y1) Then Set linha = arr(i): Exit For
    Next
    If linha Is Nothing Then Err.Raise 1000, , "Nenhuma linha reta (curva aberta de 2 nós) na seleção."
    ang = Atan2(y1 - y0, x1 - x0) * 180 / PI
    For i = 1 To n
        If arr(i).StaticID <> linha.StaticID Then
            If copia = 1 Then Set s = arr(i).Duplicate Else Set s = arr(i)
            s.RotateEx -ang, x0, y0
            s.GetBoundingBox x, y, w, h
            s.Flip cdrFlipVertical
            s.Move 0, 2 * (y0 - (y + h / 2))
            s.RotateEx ang, x0, y0
            tot = tot + 1
        End If
    Next
    Finalizar
    MsgBox tot & " objeto(s) espelhados pela linha (ângulo " & Fm(ang, 2) & "°).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_AnguloLinha()
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, ang As Double, u As Long
    If Documents.Count = 0 And Not gConfig Then Exit Sub
    If ActiveSelectionRange.Count <> 1 Then MsgBox "Selecione uma linha reta.", vbExclamation, TITULO: Exit Sub
    u = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    If EhLinha(ActiveSelectionRange(1), x0, y0, x1, y1) Then
        ang = Atan2(y1 - y0, x1 - x0) * 180 / PI
        MsgBox "Ângulo: " & Fm(ang, 3) & "°" & vbCr & _
               "Inclinação em relação à vertical: " & Fm(90 - Abs(ang), 3) & "°" & vbCr & _
               "Comprimento: " & Fm(Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2), 2) & " mm", vbInformation, TITULO
    Else
        MsgBox "O objeto não é uma linha reta de 2 nós.", vbExclamation, TITULO
    End If
    ActiveDocument.Unit = u
End Sub

Public Sub LT_ArredondarCantos()
    Dim raio As Double, col As Collection, s As Shape, so As Object, n As Long, m As String
    If Not TemSelecao() Then Exit Sub
    DlgNovo "Arredondar cantos"
    DlgNum "Raio (mm)", 3
    If Not DlgMostrar() Then Exit Sub
    raio = DV(1)
    If Not Iniciar("Arredondar cantos") Then Exit Sub
    On Error GoTo Erro
    Set col = Folhas()
    For Each s In col
        If s.Type <> cdrBitmapShape Then
            If s.Type <> cdrCurveShape Then s.ConvertToCurves
            Set so = s
            On Error Resume Next
            so.Fillet raio
            If Err.Number <> 0 Then Err.Clear: so.Fillet raio, True
            If Err.Number <> 0 Then
                m = Err.Description
                On Error GoTo Erro
                Err.Raise 1000, , "Arredondar exige CorelDRAW X7 ou superior. (" & m & ")"
            End If
            On Error GoTo Erro
            n = n + 1
        End If
    Next
    Finalizar
    MsgBox n & " objeto(s) com cantos arredondados.", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

Public Sub LT_SepararInteligente()
    Dim arr() As Shape, n As Long, i As Long, j As Long, k As Long, so As Object
    Dim pecas As ShapeRange, np As Long, p() As Shape, ar() As Double, dp() As Long, pai() As Long
    Dim px() As Variant, py() As Variant, fc() As Boolean, cpx() As Variant, cpy() As Variant
    Dim qx() As Variant, qy() As Variant, qf() As Boolean, menor As Double, sr As ShapeRange, tot As Long
    If Not Iniciar("Separar curvas") Then Exit Sub
    On Error GoTo Erro
    n = Itens(arr)
    For i = 1 To n
        If arr(i).Type = cdrTextShape Then arr(i).ConvertToCurves
        If arr(i).Type = cdrCurveShape Then
            Set so = arr(i)
            Set pecas = so.BreakApartEx
            np = pecas.Count
            ReDim p(1 To np): ReDim ar(1 To np): ReDim dp(1 To np): ReDim pai(1 To np)
            ReDim cpx(1 To np): ReDim cpy(1 To np)
            For j = 1 To np
                Set p(j) = pecas(j)
                If Poligonos(p(j), qx, qy, qf) > 0 Then
                    cpx(j) = qx(1): cpy(j) = qy(1)
                    ar(j) = Abs(AreaSinal(qx(1), qy(1)))
                End If
            Next
            For j = 1 To np
                menor = 1E+300
                For k = 1 To np
                    If k <> j And ar(k) > ar(j) And Not IsEmpty(cpx(k)) And Not IsEmpty(cpx(j)) Then
                        If DentroPoli(cpx(j)(0), cpy(j)(0), cpx(k), cpy(k)) Then
                            dp(j) = dp(j) + 1
                            If ar(k) < menor Then menor = ar(k): pai(j) = k
                        End If
                    End If
                Next
            Next
            For j = 1 To np
                If dp(j) Mod 2 = 0 Then
                    Set sr = CreateShapeRange
                    sr.Add p(j)
                    For k = 1 To np
                        If pai(k) = j And dp(k) Mod 2 = 1 Then sr.Add p(k)
                    Next
                    If sr.Count > 1 Then sr.Combine
                    tot = tot + 1
                End If
            Next
        End If
    Next
    Finalizar
    MsgBox "Separado em " & tot & " peça(s), mantendo os furos (miolo do A, O, B...).", vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

'==============================================================================
'  INFRAESTRUTURA
'==============================================================================
Private Function Iniciar(ByVal nome As String, Optional ByVal precisaSel As Boolean = True) As Boolean
    If gConfig Then
        MsgBox """" & nome & """ não tem valores para configurar.", vbInformation, TITULO
        Exit Function
    End If
    If Documents.Count = 0 Then
        MsgBox "Abra um documento primeiro.", vbExclamation, TITULO
        Exit Function
    End If
    If precisaSel Then
        If ActiveSelectionRange.Count = 0 Then
            MsgBox "Selecione um ou mais objetos.", vbExclamation, TITULO
            Exit Function
        End If
    End If
    gUnidade = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "LT: " & nome
    Optimization = True
    EventsEnabled = False
    gAtivo = True
    Iniciar = True
End Function

Private Sub Finalizar()
    If Not gAtivo Then Exit Sub
    gAtivo = False
    On Error Resume Next
    Optimization = False
    EventsEnabled = True
    ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = gUnidade
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub Falha()
    Dim m As String
    m = Err.Description
    Finalizar
    MsgBox "Não foi possível concluir:" & vbCr & m & vbCr & vbCr & _
           "Use Ctrl+Z se algo ficou pela metade.", vbCritical, TITULO
End Sub

Private Function Escolher(ByVal cab As String, opcoes As Variant) As Long
    Dim s As String, i As Long, r As String
    s = cab & vbCr & vbCr
    For i = LBound(opcoes) To UBound(opcoes)
        s = s & (i - LBound(opcoes) + 1) & " - " & opcoes(i) & vbCr
    Next
    r = Trim$(InputBox(s, TITULO, ""))
    If Len(r) > 0 And IsNumeric(r) Then Escolher = CLng(r)
End Function

Private Function SepDec() As String
    SepDec = Mid$(CStr(1.5), 2, 1)
End Function

Private Function PedirNum(ByVal msg As String, ByVal padrao As Double, ByRef valor As Double) As Boolean
    Dim r As String
    r = Trim$(InputBox(msg, TITULO, CStr(padrao)))
    If Len(r) = 0 Then Exit Function
    r = Replace(Replace(r, ",", SepDec), ".", SepDec)
    If Not IsNumeric(r) Then
        MsgBox "Valor inválido: " & r, vbExclamation, TITULO
        Exit Function
    End If
    valor = CDbl(r)
    PedirNum = True
End Function

Private Function PedirTxt(ByVal msg As String, ByVal padrao As String, ByRef valor As String) As Boolean
    Dim r As String
    r = InputBox(msg, TITULO, padrao)
    If StrPtr(r) = 0 Then Exit Function
    valor = r
    PedirTxt = True
End Function

Private Function Fm(ByVal v As Double, Optional ByVal casas As Integer = 2) As String
    Fm = FormatNumber(v, casas)
End Function

Private Function Fg(ByVal v As Double) As String
    Fg = Replace(Format$(v, "0.000"), ",", ".")
End Function

Private Function HMS(ByVal seg As Double) As String
    Dim h As Long, m As Long, s As Long
    h = Int(seg / 3600): m = Int((seg - h * 3600) / 60): s = Int(seg - h * 3600 - m * 60)
    HMS = h & "h " & Format$(m, "00") & "min " & Format$(s, "00") & "s"
End Function

Private Function FonteComercial(ByVal w As Double) As Double
    Dim v As Variant, i As Long
    v = Array(12, 24, 36, 60, 100, 150, 200, 250, 300, 400, 500)
    For i = 0 To UBound(v)
        If w <= v(i) Then FonteComercial = v(i): Exit Function
    Next
    FonteComercial = -Int(-w / 500) * 500
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2 = Atn(y / x) + PI Else Atan2 = Atn(y / x) - PI
    ElseIf y > 0 Then
        Atan2 = PI / 2
    ElseIf y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function

Private Function Camada(ByVal nome As String) As Layer
    Dim l As Layer
    If EstStr(E_CAMADAS, "1") = "0" Then Set Camada = ActiveLayer: Exit Function
    On Error Resume Next
    Set l = ActivePage.Layers(nome)
    On Error GoTo 0
    If l Is Nothing Then Set l = ActivePage.CreateLayer(nome)
    Set Camada = l
End Function

Private Sub EstiloLinha(s As Shape, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    s.Outline.Width = 0.1
    s.Outline.Color.RGBAssign r, g, b
End Sub

Private Function Texto(lay As Layer, ByVal txt As String, ByVal cx As Double, ByVal cy As Double, ByVal tam As Double) As Shape
    Dim s As Shape, x As Double, y As Double, w As Double, h As Double
    Set s = lay.CreateArtisticText(cx, cy, txt, , , EstStr(E_FONTE, "Arial"), tam)
    s.Fill.ApplyUniformFill CorCorel(CorEstilo(E_CORTXT, 5))
    s.GetBoundingBox x, y, w, h
    s.Move cx - (x + w / 2), cy - (y + h / 2)
    Set Texto = s
End Function

Private Function Itens(ByRef arr() As Shape) As Long
    Dim sr As ShapeRange, i As Long
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then Exit Function
    ReDim arr(1 To sr.Count)
    For i = 1 To sr.Count
        Set arr(i) = sr(i)
    Next
    Itens = sr.Count
End Function

Private Sub ColetarFolhas(ByVal sr As ShapeRange, ByRef col As Collection)
    Dim s As Shape
    For Each s In sr
        If s.Type = cdrGroupShape Then
            ColetarFolhas s.Shapes.All, col
        Else
            col.Add s
        End If
    Next
End Sub

Private Function Folhas() As Collection
    Dim c As New Collection
    ColetarFolhas ActiveSelectionRange, c
    Set Folhas = c
End Function

Private Sub Totais(ByRef a As Double, ByRef p As Double, ByRef e As Long, ByRef nos As Long, ByRef q As Long)
    Dim col As Collection, s As Shape
    Set col = Folhas()
    For Each s In col
        MedirForma s, a, p, e, nos
    Next
    q = col.Count
End Sub

Private Sub MedirForma(ByVal s As Shape, ByRef area As Double, ByRef perim As Double, ByRef entradas As Long, ByRef nos As Long)
    Dim so As Object, c As Object
    Dim px() As Variant, py() As Variant, fc() As Boolean, dp() As Long, n As Long, i As Long
    Set so = s
    On Error Resume Next
    Set c = so.DisplayCurve
    On Error GoTo 0
    If c Is Nothing Then Exit Sub
    perim = perim + c.Length
    entradas = entradas + c.SubPaths.Count
    nos = nos + c.Nodes.Count
    n = Poligonos(s, px, py, fc)
    If n = 0 Then Exit Sub
    Profundidades n, px, py, fc, dp
    For i = 1 To n
        If fc(i) Then
            If dp(i) Mod 2 = 0 Then
                area = area + Abs(AreaSinal(px(i), py(i)))
            Else
                area = area - Abs(AreaSinal(px(i), py(i)))
            End If
        End If
    Next
End Sub

Private Function Centroide(ByVal s As Shape, ByRef cx As Double, ByRef cy As Double) As Boolean
    Dim px() As Variant, py() As Variant, fc() As Boolean, dp() As Long, n As Long, i As Long, j As Long, m As Long
    Dim a As Double, sx As Double, sy As Double, wt As Double, cr As Double, sa As Double, sinal As Double
    n = Poligonos(s, px, py, fc)
    If n = 0 Then Exit Function
    Profundidades n, px, py, fc, dp
    For i = 1 To n
        If fc(i) Then
            m = UBound(px(i)) + 1
            a = 0: sx = 0: sy = 0
            For j = 0 To m - 1
                cr = px(i)(j) * py(i)((j + 1) Mod m) - px(i)((j + 1) Mod m) * py(i)(j)
                a = a + cr
                sx = sx + (px(i)(j) + px(i)((j + 1) Mod m)) * cr
                sy = sy + (py(i)(j) + py(i)((j + 1) Mod m)) * cr
            Next
            a = a / 2
            If Abs(a) > 0.000001 Then
                sinal = IIf(dp(i) Mod 2 = 0, 1, -1)
                ' centro do polígono = (sx/(6a), sy/(6a)); peso = |a|
                cx = cx + sinal * Abs(a) * sx / (6 * a)
                cy = cy + sinal * Abs(a) * sy / (6 * a)
                sa = sa + sinal * Abs(a)
            End If
        End If
    Next
    If Abs(sa) < 0.000001 Then Exit Function
    cx = cx / sa: cy = cy / sa
    Centroide = True
End Function

' Converte o contorno de qualquer objeto em polígonos (lista de pontos por contorno)
Private Function Poligonos(ByVal s As Shape, ByRef px() As Variant, ByRef py() As Variant, ByRef fechado() As Boolean) As Long
    Dim so As Object, c As Object, sp As Object, sg As Object
    Dim n As Long, k As Long, m As Long, xs() As Double, ys() As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim t As Double, u As Double
    Set so = s
    On Error Resume Next
    Set c = so.DisplayCurve
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    If c.SubPaths.Count = 0 Then Exit Function
    ReDim px(1 To c.SubPaths.Count): ReDim py(1 To c.SubPaths.Count): ReDim fechado(1 To c.SubPaths.Count)
    For Each sp In c.SubPaths
        n = n + 1
        m = 0
        ReDim xs(0 To 31): ReDim ys(0 To 31)
        AddPt xs, ys, m, sp.StartNode.PositionX, sp.StartNode.PositionY
        For Each sg In sp.Segments
            If sg.Type = cdrCurveSegment Then
                x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
                x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
                x1 = sg.StartingControlPointX: y1 = sg.StartingControlPointY
                x2 = sg.EndingControlPointX: y2 = sg.EndingControlPointY
                For k = 1 To AMOSTRAS
                    t = k / AMOSTRAS: u = 1 - t
                    AddPt xs, ys, m, u * u * u * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3, _
                                     u * u * u * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3
                Next
            Else
                AddPt xs, ys, m, sg.EndNode.PositionX, sg.EndNode.PositionY
            End If
        Next
        If sp.Closed And m > 2 Then
            If Abs(xs(m - 1) - xs(0)) < 0.0001 And Abs(ys(m - 1) - ys(0)) < 0.0001 Then m = m - 1
        End If
        ReDim Preserve xs(0 To m - 1): ReDim Preserve ys(0 To m - 1)
        px(n) = xs: py(n) = ys: fechado(n) = sp.Closed
    Next
    Poligonos = n
End Function

Private Sub AddPt(xs() As Double, ys() As Double, m As Long, ByVal x As Double, ByVal y As Double)
    If m > 0 Then
        If Abs(xs(m - 1) - x) < 0.00001 And Abs(ys(m - 1) - y) < 0.00001 Then Exit Sub
    End If
    If m > UBound(xs) Then
        ReDim Preserve xs(0 To UBound(xs) * 2 + 1)
        ReDim Preserve ys(0 To UBound(ys) * 2 + 1)
    End If
    xs(m) = x: ys(m) = y: m = m + 1
End Sub

Private Sub Profundidades(ByVal n As Long, px() As Variant, py() As Variant, fc() As Boolean, dp() As Long)
    Dim i As Long, j As Long
    ReDim dp(1 To n)
    For i = 1 To n
        For j = 1 To n
            If i <> j And fc(j) Then
                If DentroPoli(px(i)(0), py(i)(0), px(j), py(j)) Then dp(i) = dp(i) + 1
            End If
        Next
    Next
End Sub

Private Function AreaSinal(xs As Variant, ys As Variant) As Double
    Dim i As Long, j As Long, n As Long, a As Double
    n = UBound(xs) + 1
    For i = 0 To n - 1
        j = (i + 1) Mod n
        a = a + xs(i) * ys(j) - xs(j) * ys(i)
    Next
    AreaSinal = a / 2
End Function

Private Function DentroPoli(ByVal x As Double, ByVal y As Double, xs As Variant, ys As Variant) As Boolean
    Dim i As Long, j As Long, n As Long, r As Boolean
    n = UBound(xs) + 1
    If n < 3 Then Exit Function
    j = n - 1
    For i = 0 To n - 1
        If (ys(i) > y) <> (ys(j) > y) Then
            If x < (xs(j) - xs(i)) * (y - ys(i)) / (ys(j) - ys(i)) + xs(i) Then r = Not r
        End If
        j = i
    Next
    DentroPoli = r
End Function

Private Function DentroMaterial(ByVal x As Double, ByVal y As Double, ByVal n As Long, px() As Variant, py() As Variant, fc() As Boolean) As Boolean
    Dim i As Long, c As Long
    For i = 1 To n
        If fc(i) Then
            If DentroPoli(x, y, px(i), py(i)) Then c = c + 1
        End If
    Next
    DentroMaterial = (c Mod 2 = 1)
End Function

Private Function DistBorda(ByVal x As Double, ByVal y As Double, ByVal n As Long, px() As Variant, py() As Variant) As Double
    Dim i As Long, j As Long, m As Long, d As Double, best As Double
    best = 1E+300
    For i = 1 To n
        m = UBound(px(i)) + 1
        For j = 0 To m - 1
            d = DistSeg(x, y, px(i)(j), py(i)(j), px(i)((j + 1) Mod m), py(i)((j + 1) Mod m))
            If d < best Then best = d
        Next
    Next
    DistBorda = best
End Function

Private Function DistSeg(ByVal x As Double, ByVal y As Double, ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, L2 As Double, t As Double
    dx = bx - ax: dy = by - ay
    L2 = dx * dx + dy * dy
    If L2 > 0 Then t = ((x - ax) * dx + (y - ay) * dy) / L2
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    DistSeg = Sqr((ax + t * dx - x) ^ 2 + (ay + t * dy - y) ^ 2)
End Function

Private Function CaixaDentro(ByVal ax As Double, ByVal ay As Double, ByVal aw As Double, ByVal ah As Double, _
                             ByVal bx As Double, ByVal by As Double, ByVal bw As Double, ByVal bh As Double) As Boolean
    ' caixa A está dentro da caixa B (e é menor)
    CaixaDentro = ax >= bx - 0.001 And ay >= by - 0.001 And ax + aw <= bx + bw + 0.001 And ay + ah <= by + bh + 0.001 _
                  And aw * ah < bw * bh
End Function

Private Sub OrdenarIdx(idx() As Long, chave() As Double, ByVal n As Long)
    Dim i As Long, j As Long, v As Long
    For i = 2 To n
        v = idx(i): j = i - 1
        Do While j >= 1
            If chave(idx(j)) <= chave(v) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = v
    Next
End Sub

Private Function CriarPoligono(lay As Layer, xs() As Double, ys() As Double, ByVal n As Long, ByVal ox As Double, ByVal oy As Double) As Shape
    Dim crv As Object, sp As Object, i As Long
    Set crv = ActiveDocument.CreateCurve
    Set sp = crv.CreateSubPath(ox + xs(0), oy + ys(0))
    For i = 1 To n - 1
        sp.AppendLineSegment ox + xs(i), oy + ys(i)
    Next
    sp.Closed = True
    Set CriarPoligono = lay.CreateCurve(crv)
End Function

Private Function EhLinha(ByVal s As Shape, ByRef x0 As Double, ByRef y0 As Double, ByRef x1 As Double, ByRef y1 As Double) As Boolean
    Dim so As Object, c As Object
    Set so = s
    On Error Resume Next
    Set c = so.DisplayCurve
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    If c.SubPaths.Count <> 1 Then Exit Function
    If c.SubPaths(1).Closed Then Exit Function
    If c.Nodes.Count <> 2 Then Exit Function
    x0 = c.SubPaths(1).StartNode.PositionX: y0 = c.SubPaths(1).StartNode.PositionY
    x1 = c.SubPaths(1).EndNode.PositionX: y1 = c.SubPaths(1).EndNode.PositionY
    EhLinha = True
End Function

' Duplica o objeto e transforma em uma única curva sólida (para contornar)
Private Function BaseSolida(ByVal s As Shape) As Shape
    Dim d As Shape, sr As ShapeRange, ok As ShapeRange, x As Shape, u As Shape, i As Long, r As Shape
    Dim bx As Double, by As Double, bw As Double, bh As Double, so As Object
    Set d = s.Duplicate
    If d.Type = cdrGroupShape Then
        Set so = d
        Set sr = so.UngroupAllEx
    Else
        Set sr = CreateShapeRange: sr.Add d
    End If
    Set ok = CreateShapeRange
    For Each x In sr
        If x.Type = cdrBitmapShape Then
            x.GetBoundingBox bx, by, bw, bh
            Set r = x.Layer.CreateRectangle2(bx, by, bw, bh)
            x.Delete
            ok.Add r
        Else
            If x.Type <> cdrCurveShape Then x.ConvertToCurves
            ok.Add x
        End If
    Next
    Set u = ok(1)
    For i = 2 To ok.Count
        Set u = u.Weld(ok(i))
    Next
    Set BaseSolida = u
End Function

' Junta toda a seleção (cópia) em uma curva combinada
Private Function BaseCombinada(ByVal sr As ShapeRange) As Shape
    Dim d As ShapeRange, col As New Collection, s As Shape, ok As ShapeRange, so As Object
    Set d = sr.Duplicate
    ColetarFolhas d, col
    Set ok = CreateShapeRange
    For Each s In col
        If s.Type <> cdrBitmapShape Then
            If s.Type <> cdrCurveShape Then s.ConvertToCurves
            ok.Add s
        End If
    Next
    If ok.Count = 1 Then
        Set BaseCombinada = ok(1)
    Else
        Set so = ok
        Set BaseCombinada = so.Combine
    End If
End Function

' Cria contorno (contour) e devolve as curvas resultantes já separadas
Private Function Contornar(ByVal base As Shape, ByVal dirc As Long, ByVal off As Double) As ShapeRange
    Dim so As Object, ef As Object, sr As Object, x As Shape, res As ShapeRange, so2 As Object
    Set res = CreateShapeRange
    base.Name = "LT_BASE"
    Set so = base
    Set ef = so.CreateContour(dirc, off, 1)
    Set sr = ef.Separate
    For Each x In sr
        If x.Name <> "LT_BASE" Then
            If x.Type = cdrGroupShape Then
                Set so2 = x
                res.AddRange so2.UngroupAllEx
            Else
                res.Add x
            End If
        End If
    Next
    Set Contornar = res
End Function

'==============================================================================
'  JANELAS (usadas pelo formulário frmLT)
'==============================================================================
Private Sub DlgNovo(ByVal tit As String, Optional ByVal dica As String = "")
    gTit = tit: gDica = dica: gN = 0: gOK = False
End Sub

Private Sub AddCampo(ByVal tipo As Long, ByVal rot As String, ByVal valor As String, ByVal opc As String)
    If gN >= 16 Then Exit Sub
    gN = gN + 1
    gTipo(gN) = tipo: gRot(gN) = rot: gVal(gN) = valor: gOpc(gN) = opc
End Sub

Private Sub DlgNum(ByVal rot As String, ByVal padrao As Double)
    AddCampo 1, rot, CStr(padrao), ""
End Sub

Private Sub DlgSim(ByVal rot As String, ByVal padrao As Double)
    AddCampo 2, rot, IIf(padrao <> 0, "1", "0"), ""
End Sub

Private Sub DlgLista(ByVal rot As String, ByVal itens As String, ByVal padrao As Long)
    AddCampo 3, rot, CStr(padrao), itens
End Sub

Private Sub DlgTexto(ByVal rot As String, ByVal padrao As String)
    AddCampo 4, rot, padrao, ""
End Sub

Private Function DlgMostrar() As Boolean
    Dim i As Long, s As String, partes As Variant
    For i = 1 To gN
        gFab(i) = gVal(i)
        s = GetSetting("LTTools", gTit, gRot(i), "")
        If Len(s) > 0 Then gVal(i) = s
        gUlt(i) = gVal(i)
    Next
    ' padrão definido pelo usuário tem prioridade ao abrir
    s = GetSetting("LTTools", "Pre_" & gTit, "_padrao", "")
    If Len(s) > 0 Then
        partes = Split(s, vbTab)
        For i = 1 To gN
            If i - 1 <= UBound(partes) Then gVal(i) = partes(i - 1)
        Next
    End If
    gModo = 1: gOK = False
    If Not MostrarForm() Then Exit Function
    If gOK Then
        If gConfig Then
            SaveSetting "LTTools", "Pre_" & gTit, "_padrao", JuntarValores()
            MsgBox "Valores padrão salvos para """ & gTit & """.", vbInformation, TITULO
            Exit Function                      ' em modo configuração não executa
        End If
        For i = 1 To gN
            SaveSetting "LTTools", gTit, gRot(i), gVal(i)
        Next
    End If
    DlgMostrar = gOK
End Function

Private Function JuntarValores() As String
    Dim i As Long, s As String
    For i = 1 To gN
        s = s & IIf(i > 1, vbTab, "") & gVal(i)
    Next
    JuntarValores = s
End Function

Private Function AbrirMenu(ByVal tit As String, itens As Variant, icones As Variant) As Long
    Dim i As Long
    gTit = tit: gMenuN = 0: gMenuEscolha = 0
    For i = LBound(itens) To UBound(itens)
        If gMenuN < 16 Then
            gMenuN = gMenuN + 1
            gMenuTxt(gMenuN) = itens(i)
            gMenuIco(gMenuN) = PastaIcones() & "m_" & icones(i) & ".ico"
        End If
    Next
    gModo = 2
    If Not MostrarForm() Then
        AbrirMenu = Escolher(tit, itens)       ' sem o formulário: usa a lista simples
        Exit Function
    End If
    AbrirMenu = gMenuEscolha
End Function

Private Function MostrarForm() As Boolean
    Dim f As Object
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmLT")
    If f Is Nothing Then
        If gModo = 1 Then
            MsgBox "Não encontrei o formulário ""frmLT"" no GlobalMacros." & vbCr & _
                   "Veja o passo 3 do LEIA-ME (inserir o UserForm).", vbCritical, TITULO
        End If
        Exit Function
    End If
    On Error GoTo 0
    f.Show
    On Error Resume Next
    Unload f
    MostrarForm = True
End Function

Private Function DV(ByVal i As Long) As Double
    Dim r As String
    r = Replace(Replace(Trim$(gVal(i)), ",", SepDec), ".", SepDec)
    If Len(r) > 0 And IsNumeric(r) Then DV = CDbl(r)
End Function

Private Function DT(ByVal i As Long) As String
    DT = gVal(i)
End Function

Public Function PastaIcones() As String
    Dim p As String, c As Variant
    p = GetSetting("LTTools", "Config", "Icones", "")
    If Len(p) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        If Len(Dir$(p & "logo.ico")) > 0 Then PastaIcones = p: Exit Function
    End If
    ' procura nos lugares mais comuns (inclusive quando o ZIP cria pasta dupla)
    For Each c In Array(Environ$("USERPROFILE") & "\Documents\Corel\Corel Content\LTTools\icones\", Environ$("USERPROFILE") & "\Documents\LTTools\icones\", "C:\LTTools\icones\", "C:\LTTools\LTTools\icones\", _
                        Environ$("USERPROFILE") & "\Downloads\LTTools\icones\", _
                        Environ$("USERPROFILE") & "\Downloads\LTTools\LTTools\icones\", _
                        Environ$("USERPROFILE") & "\Desktop\LTTools\icones\", _
                        Environ$("USERPROFILE") & "\Desktop\LTTools\LTTools\icones\")
        If Len(Dir$(c & "logo.ico")) > 0 Then
            SaveSetting "LTTools", "Config", "Icones", c
            PastaIcones = c
            Exit Function
        End If
    Next
    PastaIcones = "C:\LTTools\icones\"
End Function

Private Function TemSelecao() As Boolean
    If gConfig Then TemSelecao = True: Exit Function
    If Documents.Count = 0 Then
        MsgBox "Abra um documento primeiro.", vbExclamation, TITULO
        Exit Function
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Selecione um ou mais objetos antes de usar esta ferramenta.", vbExclamation, TITULO
        Exit Function
    End If
    TemSelecao = True
End Function

'==============================================================================
'  ESTILOS E VALORES PADRÃO
'==============================================================================
Public Sub LT_Estilos()
    DlgNovo "Estilos", "Cores, linhas e fonte que as ferramentas usam ao criar objetos."
    DlgLista E_CORTE, CORES, 1
    DlgNum E_ESPCORTE, 0.1
    DlgLista E_FURO, CORES, 2
    DlgLista E_MARCA, CORES, 3
    DlgNum E_ESPMARCA, 0.15
    DlgTexto E_FONTE, "Arial"
    DlgLista E_CORTXT, CORES, 5
    DlgLista E_LED, CORES, 7
    DlgSim E_CAMADAS, 1
    If DlgMostrar() Then MsgBox "Estilos salvos. Valem para as próximas ferramentas que você usar.", vbInformation, TITULO
End Sub

Public Sub LT_Configurar()
    Dim r As Long
    r = AbrirMenu("Valores padrão - escolha a categoria", _
        Array("Medição e orçamento", "Encaixe e otimização", "Plotter, laser e CNC", _
              "Letra caixa, LED e neon", "Geradores", "Edição de objetos"), _
        Array("cat_medicao", "cat_encaixe", "cat_corte", "cat_letracaixa", "cat_geradores", "cat_edicao"))
    If r = 0 Then Exit Sub
    gConfig = True
    On Error GoTo Fim
    Select Case r
        Case 1: LT_Medicao
        Case 2: LT_Encaixe
        Case 3: LT_Corte
        Case 4: LT_LetraCaixaLED
        Case 5: LT_Geradores
        Case 6: LT_Edicao
    End Select
Fim:
    gConfig = False
End Sub

Private Function EstStr(ByVal rot As String, ByVal padrao As String) As String
    EstStr = GetSetting("LTTools", "Estilos", rot, padrao)
    If Len(EstStr) = 0 Then EstStr = padrao
End Function

Private Function EstNum(ByVal rot As String, ByVal padrao As Double) As Double
    Dim r As String
    r = Replace(Replace(EstStr(rot, ""), ",", SepDec), ".", SepDec)
    If Len(r) > 0 And IsNumeric(r) Then EstNum = CDbl(r) Else EstNum = padrao
End Function

Private Function CorRGB(ByVal idx As Long) As Long
    Select Case idx
        Case 1: CorRGB = RGB(255, 0, 255)      ' magenta
        Case 2: CorRGB = RGB(255, 0, 0)        ' vermelho
        Case 3: CorRGB = RGB(0, 90, 255)       ' azul
        Case 4: CorRGB = RGB(0, 170, 60)       ' verde
        Case 5: CorRGB = RGB(0, 0, 0)          ' preto
        Case 6: CorRGB = RGB(0, 200, 255)      ' ciano
        Case 7: CorRGB = RGB(255, 200, 0)      ' amarelo
        Case 8: CorRGB = RGB(255, 130, 0)      ' laranja
        Case 9: CorRGB = RGB(255, 255, 255)    ' branco
        Case Else: CorRGB = RGB(130, 130, 130) ' cinza
    End Select
End Function

Private Function CorEstilo(ByVal rot As String, ByVal padraoIdx As Long) As Long
    CorEstilo = CorRGB(Val(EstStr(rot, CStr(padraoIdx))))
End Function

Private Function CorCorel(ByVal c As Long) As Color
    Set CorCorel = CreateRGBColor(c Mod 256, (c \ 256) Mod 256, c \ 65536)
End Function

Private Sub Estilo(s As Shape, ByVal tipo As String)
    Dim c As Long, w As Double
    Select Case tipo
        Case "furo": c = CorEstilo(E_FURO, 2): w = EstNum(E_ESPCORTE, 0.1)
        Case "marca": c = CorEstilo(E_MARCA, 3): w = EstNum(E_ESPMARCA, 0.15)
        Case Else: c = CorEstilo(E_CORTE, 1): w = EstNum(E_ESPCORTE, 0.1)
    End Select
    s.Outline.Width = w
    s.Outline.Color.RGBAssign c Mod 256, (c \ 256) Mod 256, c \ 65536
End Sub

' nome do projeto (.gms) onde o LT Tools está instalado: LTPlugin (instalador) ou GlobalMacros
Public Function ProjetoLT() As String
    Dim app As Object, p As Object, c As Object
    ProjetoLT = "GlobalMacros"
    On Error Resume Next
    Set app = Application
    For Each p In app.VBE.VBProjects
        If LCase$(p.Name) = "ltplugin" Then
            Set c = Nothing
            Set c = p.VBComponents("LTTools")
            If Not c Is Nothing Then ProjetoLT = p.Name: Exit Function
        End If
    Next
End Function

'==============================================================================
'  ENCAIXE PELA FORMA (giro livre, peças dentro de furos, visualização)
'  Cada peça vira uma grade de quadradinhos (raster) em cada ângulo; o encaixe
'  procura a posição mais baixa (ou mais à esquerda) onde ela cabe na chapa.
'==============================================================================
Public Sub LT_EncaixeForma()
    Dim ok As Boolean, u As Long
    If Not TemSelecao() Then Exit Sub
    If Not gConfig Then
        If Not NzColetar() Then Exit Sub
    End If
    DlgNovo "Encaixe pela forma", "Veja o resultado à esquerda. Mude as opções e clique em Visualizar; OK aplica no desenho."
    DlgNum "Largura da chapa (mm)", 1000
    DlgNum "Altura da chapa (mm)", 370
    DlgSim "Usar o tamanho da página ativa", 0
    DlgNum "Margem da chapa (mm)", 0
    DlgNum "Distância mínima entre peças (mm)", 2
    DlgSim "Permitir peças dentro de furos", 1
    DlgLista "Girar peças", "Não girar|90°|45°|30°|15°|10°", 4
    DlgLista "Encaixar para", "Baixo (melhor)|Esquerda", 1
    DlgLista "Precisão", "Baixa (rápido)|Média|Alta|Muito alta", 2
    DlgLista "Forma usada", "Forma real da peça|Retângulo da peça", 1
    DlgSim "Buscar melhor resultado (testa várias ordens)", 0
    DlgNum "Tempo de busca (s)", 10
    gPreview = Not gConfig
    gPrevArq = "": gPrevInfo = ""
    ok = DlgMostrar()
    gPreview = False
    If Not ok Then Exit Sub
    If NzAssinatura() <> nzAssin Then NzCalcular
    If nzNS = 0 Then
        MsgBox "Nenhuma peça coube na chapa.", vbExclamation, TITULO
        Exit Sub
    End If
    If Not Iniciar("Encaixe pela forma") Then Exit Sub
    On Error GoTo Erro
    NzAplicar
    Finalizar
    MsgBox gPrevInfo, vbInformation, TITULO
    Exit Sub
Erro:
    Falha
End Sub

' chamado pela janela (botão Visualizar)
Public Sub NzPrevia()
    On Error GoTo Erro
    If NzAssinatura() <> nzAssin Then NzCalcular
    NzDesenharPrevia
    Exit Sub
Erro:
    gPrevInfo = "Erro ao calcular: " & Err.Description
End Sub

' guarda os contornos das peças selecionadas (em mm)
Private Function NzColetar() As Boolean
    Dim u As Long, sr As ShapeRange, i As Long, col As Collection, s As Shape
    Dim px() As Variant, py() As Variant, fc() As Boolean, n As Long, j As Long
    Dim ax() As Variant, ay() As Variant, af() As Variant, m As Long
    Dim xmn As Double, xmx As Double, ymn As Double, ymx As Double, k As Long
    u = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    On Error GoTo Erro
    Set sr = ActiveSelectionRange
    nzN = sr.Count
    ReDim nzShp(1 To nzN): ReDim nzPX(1 To nzN): ReDim nzPY(1 To nzN): ReDim nzPF(1 To nzN)
    ReDim nzCX(1 To nzN): ReDim nzCY(1 To nzN)
    sr.GetBoundingBox nzSX, nzSY, nzSW, nzSH
    nzPagW = ActivePage.SizeWidth: nzPagH = ActivePage.SizeHeight
    For i = 1 To nzN
        Set nzShp(i) = sr(i)
        Set col = New Collection
        If sr(i).Type = cdrGroupShape Then ColetarFolhas sr(i).Shapes.All, col Else col.Add sr(i)
        m = 0
        ReDim ax(0 To 0): ReDim ay(0 To 0): ReDim af(0 To 0)
        xmn = 1E+300: ymn = 1E+300: xmx = -1E+300: ymx = -1E+300
        For Each s In col
            n = Poligonos(s, px, py, fc)
            For j = 1 To n
                m = m + 1
                ReDim Preserve ax(0 To m - 1): ReDim Preserve ay(0 To m - 1): ReDim Preserve af(0 To m - 1)
                ax(m - 1) = px(j): ay(m - 1) = py(j): af(m - 1) = fc(j)
                For k = 0 To UBound(px(j))
                    If px(j)(k) < xmn Then xmn = px(j)(k)
                    If px(j)(k) > xmx Then xmx = px(j)(k)
                    If py(j)(k) < ymn Then ymn = py(j)(k)
                    If py(j)(k) > ymx Then ymx = py(j)(k)
                Next
            Next
        Next
        If m = 0 Then        ' sem contorno (ex.: bitmap): usa o retângulo
            sr(i).GetBoundingBox xmn, ymn, xmx, ymx
            xmx = xmn + xmx: ymx = ymn + ymx
            ax(0) = Array(xmn, xmx, xmx, xmn): ay(0) = Array(ymn, ymn, ymx, ymx): af(0) = True
        End If
        nzPX(i) = ax: nzPY(i) = ay: nzPF(i) = af
        nzCX(i) = (xmn + xmx) / 2: nzCY(i) = (ymn + ymx) / 2
    Next
    ActiveDocument.Unit = u
    nzAssin = ""
    NzColetar = True
    Exit Function
Erro:
    ActiveDocument.Unit = u
    MsgBox "Não consegui ler as peças: " & Err.Description, vbExclamation, TITULO
End Function

Private Function NzAssinatura() As String
    Dim i As Long, s As String
    For i = 1 To gN
        s = s & gVal(i) & "|"
    Next
    NzAssinatura = s
End Function

Private Sub NzCalcular()
    Dim larg As Double, alt As Double, dist As Double, passo As Double, fator As Long, a As Long
    Dim dentro As Boolean, retang As Boolean, esquerda As Boolean, buscar As Boolean, tempo As Double
    Dim t0 As Single, tentativa As Long, ordem() As Long, chave() As Double, i As Long, j As Long, tmp As Long
    Dim melhorS() As Long, melhorAI() As Long, melhorX() As Long, melhorY() As Long, melhorNS As Long, melhorAlt As Long
    Dim altUsada As Long, melhorFora As Long
    larg = DV(1): alt = DV(2)
    If DV(3) = 1 Then larg = nzPagW: alt = nzPagH
    nzMarg = DV(4): dist = DV(5): dentro = (DV(6) = 1)
    Select Case DV(7)
        Case 1: passo = 0
        Case 2: passo = 90
        Case 3: passo = 45
        Case 4: passo = 30
        Case 5: passo = 15
        Case Else: passo = 10
    End Select
    esquerda = (DV(8) = 2)
    Select Case DV(9)
        Case 1: fator = 200
        Case 2: fator = 400
        Case 3: fator = 700
        Case Else: fator = 1100
    End Select
    retang = (DV(10) = 2)
    buscar = (DV(11) = 1): tempo = DV(12)
    nzSWch = larg: nzSHch = alt
    larg = larg - 2 * nzMarg: alt = alt - 2 * nzMarg
    If larg <= 0 Or alt <= 0 Then nzNS = 0: gPrevInfo = "Chapa menor que a margem.": nzAssin = NzAssinatura(): Exit Sub
    nzRes = IIf(larg > alt, larg, alt) / fator
    nzGW = Int(larg / nzRes): nzGH = Int(alt / nzRes)
    nzK = 0
    If dist > 0 Then nzK = -Int(-dist / nzRes)
    If passo = 0 Then
        nzNA = 1: ReDim nzAng(1 To 1): nzAng(1) = 0
    Else
        nzNA = Int(360 / passo + 0.5)
        ReDim nzAng(1 To nzNA)
        For a = 1 To nzNA
            nzAng(a) = (a - 1) * passo
        Next
    End If
    ' rasters
    nRu = 0: nRd = 0
    ReDim ruR(1 To 1024): ReDim ruA(1 To 1024): ReDim ruB(1 To 1024)
    ReDim rdR(1 To 1024): ReDim rdA(1 To 1024): ReDim rdB(1 To 1024)
    ReDim rsW(1 To nzN * nzNA): ReDim rsH(1 To nzN * nzNA): ReDim rsXm(1 To nzN * nzNA): ReDim rsYm(1 To nzN * nzNA)
    ReDim rsOk(1 To nzN * nzNA): ReDim rsArea(1 To nzN * nzNA): ReDim rsKW(1 To nzN * nzNA)
    ReDim rsR0(1 To nzN * nzNA): ReDim rsR1(1 To nzN * nzNA): ReDim rsD0(1 To nzN * nzNA): ReDim rsD1(1 To nzN * nzNA)
    For i = 1 To nzN
        For a = 1 To nzNA
            NzRaster i, a, (i - 1) * nzNA + a, dentro, retang
        Next
        If i Mod 5 = 0 Then DoEvents
    Next
    ' ordem: maiores primeiro
    ReDim ordem(1 To nzN): ReDim chave(1 To nzN)
    For i = 1 To nzN
        ordem(i) = i
        chave(i) = -rsArea((i - 1) * nzNA + 1)
    Next
    OrdenarIdx ordem, chave, nzN
    t0 = Timer
    tentativa = 0
    melhorNS = 0
    Do
        tentativa = tentativa + 1
        NzColocar ordem, esquerda, altUsada
        If melhorNS = 0 Or (nzNS > 0 And (nzNS < melhorNS Or (nzNS = melhorNS And altUsada < melhorAlt))) Then
            melhorNS = nzNS: melhorAlt = altUsada: melhorFora = nzFora
            melhorS = nzS: melhorAI = nzAI: melhorX = nzX: melhorY = nzY
        End If
        If Not buscar Then Exit Do
        If Timer - t0 > tempo Or Timer < t0 Then Exit Do
        ' próxima tentativa: embaralha um pouco a ordem
        Randomize
        For i = 1 To nzN
            j = Int(Rnd * nzN) + 1
            If Rnd < 0.35 Then tmp = ordem(i): ordem(i) = ordem(j): ordem(j) = tmp
        Next
    Loop
    nzNS = melhorNS: nzFora = melhorFora
    If melhorNS > 0 Then nzS = melhorS: nzAI = melhorAI: nzX = melhorX: nzY = melhorY
    nzAssin = NzAssinatura()
    gPrevInfo = "Chapas: " & nzNS & "   |   Aproveitamento: " & _
                IIf(nzNS > 0, Fm(100 * nzAreaPecas / (nzNS * nzSWch * nzSHch), 1) & "%", "-") & _
                "   |   Tentativas: " & tentativa & _
                IIf(nzFora > 0, "   |   " & nzFora & " peça(s) não cabem", "")
End Sub

' rasteriza a peça i girada no ângulo a (id = índice do raster)
Private Sub NzRaster(ByVal i As Long, ByVal a As Long, ByVal id As Long, ByVal dentro As Boolean, ByVal retang As Boolean)
    Dim ca As Double, sa As Double, np As Long, p As Long, q As Long, m As Long
    Dim rx() As Variant, ry() As Variant, usa() As Boolean, xs() As Double, ys() As Double
    Dim xmn As Double, ymn As Double, xmx As Double, ymx As Double, x As Double, y As Double
    Dim W As Long, H As Long, g() As Byte, cnt() As Long, ofs() As Long, cur() As Long, xi() As Double, tot As Long
    Dim r As Long, r0 As Long, r1 As Long, c As Long, c0 As Long, c1 As Long, k As Long, yc As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, t As Double, st As Long, L As Double, v As Double
    Dim W2 As Long, H2 As Long, d() As Byte, rr As Long, cc As Long, area As Long, ang As Double
    ang = nzAng(a) * PI / 180
    ca = Cos(ang): sa = Sin(ang)
    np = UBound(nzPX(i)) + 1
    ReDim usa(0 To np - 1)
    For p = 0 To np - 1
        usa(p) = True
    Next
    If Not dentro Then          ' sem furos: usa só os contornos externos
        For p = 0 To np - 1
            If nzPF(i)(p) Then
                For q = 0 To np - 1
                    If q <> p And nzPF(i)(q) Then
                        If DentroPoli(nzPX(i)(p)(0), nzPY(i)(p)(0), nzPX(i)(q), nzPY(i)(q)) Then usa(p) = False: Exit For
                    End If
                Next
            End If
        Next
    End If
    ' gira os pontos em volta do centro da peça
    ReDim rx(0 To np - 1): ReDim ry(0 To np - 1)
    xmn = 1E+300: ymn = 1E+300: xmx = -1E+300: ymx = -1E+300
    For p = 0 To np - 1
        m = UBound(nzPX(i)(p))
        ReDim xs(0 To m): ReDim ys(0 To m)
        For q = 0 To m
            x = nzPX(i)(p)(q) - nzCX(i): y = nzPY(i)(p)(q) - nzCY(i)
            xs(q) = nzCX(i) + x * ca - y * sa
            ys(q) = nzCY(i) + x * sa + y * ca
            If xs(q) < xmn Then xmn = xs(q)
            If xs(q) > xmx Then xmx = xs(q)
            If ys(q) < ymn Then ymn = ys(q)
            If ys(q) > ymx Then ymx = ys(q)
        Next
        rx(p) = xs: ry(p) = ys
    Next
    W = Int((xmx - xmn) / nzRes) + 1: H = Int((ymx - ymn) / nzRes) + 1
    rsW(id) = W: rsH(id) = H: rsXm(id) = xmn: rsYm(id) = ymn
    rsOk(id) = (W <= nzGW And H <= nzGH)
    rsR0(id) = nRu + 1: rsR1(id) = nRu: rsD0(id) = nRd + 1: rsD1(id) = nRd
    If Not rsOk(id) Then Exit Sub
    ReDim g(0 To W * H - 1)
    If retang Then
        For k = 0 To W * H - 1
            g(k) = 1
        Next
    Else
        ' 1) preenchimento par-ímpar pelo centro das células
        ReDim cnt(0 To H - 1)
        For p = 0 To np - 1
            If usa(p) And nzPF(i)(p) Then
                m = UBound(rx(p))
                For q = 0 To m
                    y1 = ry(p)(q): y2 = ry(p)((q + 1) Mod (m + 1))
                    If y1 <> y2 Then
                        If y1 > y2 Then v = y1: y1 = y2: y2 = v
                        r0 = -Int(-((y1 - ymn) / nzRes - 0.5)): r1 = -Int(-((y2 - ymn) / nzRes - 0.5)) - 1
                        If r0 < 0 Then r0 = 0
                        If r1 > H - 1 Then r1 = H - 1
                        For r = r0 To r1
                            cnt(r) = cnt(r) + 1
                        Next
                    End If
                Next
            End If
        Next
        ReDim ofs(0 To H): ReDim cur(0 To H - 1)
        For r = 0 To H - 1
            ofs(r + 1) = ofs(r) + cnt(r): cur(r) = ofs(r)
        Next
        tot = ofs(H)
        If tot > 0 Then
            ReDim xi(0 To tot - 1)
            For p = 0 To np - 1
                If usa(p) And nzPF(i)(p) Then
                    m = UBound(rx(p))
                    For q = 0 To m
                        x1 = rx(p)(q): y1 = ry(p)(q)
                        x2 = rx(p)((q + 1) Mod (m + 1)): y2 = ry(p)((q + 1) Mod (m + 1))
                        If y1 <> y2 Then
                            If y1 > y2 Then v = y1: y1 = y2: y2 = v: v = x1: x1 = x2: x2 = v
                            r0 = -Int(-((y1 - ymn) / nzRes - 0.5)): r1 = -Int(-((y2 - ymn) / nzRes - 0.5)) - 1
                            If r0 < 0 Then r0 = 0
                            If r1 > H - 1 Then r1 = H - 1
                            For r = r0 To r1
                                yc = ymn + (r + 0.5) * nzRes
                                xi(cur(r)) = x1 + (yc - y1) * (x2 - x1) / (y2 - y1)
                                cur(r) = cur(r) + 1
                            Next
                        End If
                    Next
                End If
            Next
            For r = 0 To H - 1
                ' ordena as interseções da linha (inserção)
                For k = ofs(r) + 1 To ofs(r + 1) - 1
                    v = xi(k): q = k - 1
                    Do While q >= ofs(r)
                        If xi(q) <= v Then Exit Do
                        xi(q + 1) = xi(q): q = q - 1
                    Loop
                    xi(q + 1) = v
                Next
                For k = ofs(r) To ofs(r + 1) - 2 Step 2
                    c0 = Int((xi(k) - xmn) / nzRes + 0.5): c1 = Int((xi(k + 1) - xmn) / nzRes - 0.5)
                    If c0 < 0 Then c0 = 0
                    If c1 > W - 1 Then c1 = W - 1
                    For c = c0 To c1
                        g(r * W + c) = 1
                    Next
                Next
            Next
        End If
        ' 2) marca as células que a linha do contorno atravessa (traços finos)
        For p = 0 To np - 1
            If usa(p) Then
                m = UBound(rx(p))
                For q = 0 To IIf(nzPF(i)(p), m, m - 1)
                    x1 = rx(p)(q): y1 = ry(p)(q)
                    x2 = rx(p)((q + 1) Mod (m + 1)): y2 = ry(p)((q + 1) Mod (m + 1))
                    L = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
                    st = Int(L / (nzRes * 0.5)) + 1
                    For k = 0 To st
                        t = k / st
                        c = Int((x1 + (x2 - x1) * t - xmn) / nzRes): r = Int((y1 + (y2 - y1) * t - ymn) / nzRes)
                        If c > W - 1 Then c = W - 1
                        If r > H - 1 Then r = H - 1
                        If c < 0 Then c = 0
                        If r < 0 Then r = 0
                        g(r * W + c) = 1
                    Next
                Next
            End If
        Next
    End If
    For k = 0 To W * H - 1
        area = area + g(k)
    Next
    rsArea(id) = area
    NzRuns g, W, H, 0, ruR, ruA, ruB, nRu
    rsR1(id) = nRu
    ' versão "engordada" pela distância mínima (marca na chapa)
    If nzK > 0 Then
        W2 = W + 2 * nzK: H2 = H + 2 * nzK
        ReDim d(0 To W2 * H2 - 1)
        For r = 0 To H - 1
            For c = 0 To W - 1
                If g(r * W + c) Then
                    For rr = r To r + 2 * nzK
                        For cc = c To c + 2 * nzK
                            d(rr * W2 + cc) = 1
                        Next
                    Next
                End If
            Next
        Next
        NzRuns d, W2, H2, nzK, rdR, rdA, rdB, nRd
    Else
        NzRuns g, W, H, 0, rdR, rdA, rdB, nRd
    End If
    rsD1(id) = nRd
End Sub

' transforma a grade em faixas (linha, coluna inicial, coluna final), deslocadas por -k
Private Sub NzRuns(g() As Byte, ByVal W As Long, ByVal H As Long, ByVal k As Long, aR() As Long, aA() As Long, aB() As Long, n As Long)
    Dim r As Long, c As Long, a As Long
    For r = 0 To H - 1
        c = 0
        Do While c < W
            If g(r * W + c) Then
                a = c
                Do While c < W
                    If g(r * W + c) = 0 Then Exit Do
                    c = c + 1
                Loop
                n = n + 1
                If n > UBound(aR) Then
                    ReDim Preserve aR(1 To n * 2): ReDim Preserve aA(1 To n * 2): ReDim Preserve aB(1 To n * 2)
                End If
                aR(n) = r - k: aA(n) = a - k: aB(n) = c - 1 - k
            Else
                c = c + 1
            End If
        Loop
    Next
End Sub

' coloca as peças na ordem dada; resultado em nzS/nzAI/nzX/nzY
Private Sub NzColocar(ordem() As Long, ByVal esquerda As Boolean, ByRef altUsada As Long)
    Dim occ() As Byte, pre() As Long, ult() As Long, GW1 As Long, tamO As Long, tamP As Long, nx As Long, ultC As Long
    Dim k As Long, i As Long, a As Long, id As Long, s As Long, x As Long, y As Long, q As Long
    Dim W As Long, H As Long, o As Long, base As Long, cabe As Boolean, achou As Boolean
    Dim melhor As Double, sc As Double, bA As Long, bX As Long, bY As Long, lim As Long
    Dim r As Long, c As Long, c0 As Long, c1 As Long, acc As Long, yy As Long, xa As Long, ya As Long
    GW1 = nzGW + 1
    tamO = nzGW * nzGH: tamP = GW1 * nzGH
    ReDim nzS(1 To nzN): ReDim nzAI(1 To nzN): ReDim nzX(1 To nzN): ReDim nzY(1 To nzN)
    nzNS = 0: nzFora = 0: nzAreaPecas = 0: altUsada = 0
    ReDim occ(0 To 0): ReDim pre(0 To 0): ReDim ult(0 To 0)
    For k = 1 To nzN
        i = ordem(k)
        nzS(i) = 0
        cabe = False
        For a = 1 To nzNA
            If rsOk((i - 1) * nzNA + a) Then cabe = True: Exit For
        Next
        If Not cabe Then
            nzFora = nzFora + 1
        Else
            s = 0: bA = 0
            Do
                If s >= nzNS Then            ' nova chapa
                    nzNS = nzNS + 1
                    ReDim Preserve occ(0 To nzNS * tamO - 1)
                    ReDim Preserve pre(0 To nzNS * tamP - 1)
                    ReDim Preserve ult(0 To nzNS * tamO - 1)
                    For c = (nzNS - 1) * tamO To nzNS * tamO - 1
                        ult(c) = -1
                    Next
                End If
                base = s * tamP
                melhor = 1E+300
                For a = 1 To nzNA
                    id = (i - 1) * nzNA + a
                    If rsOk(id) Then
                        W = rsW(id): H = rsH(id)
                        achou = False
                        If Not esquerda Then
                            lim = nzGH - H
                            If melhor < 1E+299 Then
                                If bY + rsH((i - 1) * nzNA + bA) - H < lim Then lim = bY + rsH((i - 1) * nzNA + bA) - H
                            End If
                            For y = 0 To lim
                                x = 0
                                Do While x <= nzGW - W
                                    cabe = True
                                    For q = rsR0(id) To rsR1(id)
                                        o = base + (y + ruR(q)) * GW1 + x
                                        If pre(o + ruB(q) + 1) - pre(o + ruA(q)) > 0 Then
                                            ' pula direto para depois do último ponto ocupado
                                            nx = ult(s * tamO + (y + ruR(q)) * nzGW + x + ruB(q)) - ruA(q) + 1
                                            cabe = False: Exit For
                                        End If
                                    Next
                                    If cabe Then
                                        sc = CDbl(y + H) * 1000000# + x
                                        If sc < melhor Then melhor = sc: bA = a: bX = x: bY = y
                                        achou = True: Exit Do
                                    End If
                                    If nx <= x Then nx = x + 1
                                    x = nx
                                Loop
                                If achou Then Exit For
                            Next
                        Else
                            lim = nzGW - W
                            If melhor < 1E+299 Then
                                If bX + rsW((i - 1) * nzNA + bA) - W < lim Then lim = bX + rsW((i - 1) * nzNA + bA) - W
                            End If
                            For x = 0 To lim
                                For y = 0 To nzGH - H
                                    cabe = True
                                    For q = rsR0(id) To rsR1(id)
                                        o = base + (y + ruR(q)) * GW1 + x
                                        If pre(o + ruB(q) + 1) - pre(o + ruA(q)) > 0 Then cabe = False: Exit For
                                    Next
                                    If cabe Then
                                        sc = CDbl(x + W) * 1000000# + y
                                        If sc < melhor Then melhor = sc: bA = a: bX = x: bY = y
                                        achou = True: Exit For
                                    End If
                                Next
                                If achou Then Exit For
                            Next
                        End If
                    End If
                Next
                If bA > 0 Then Exit Do
                s = s + 1
            Loop
            ' marca a peça (engordada) na chapa e refaz as somas das linhas tocadas
            id = (i - 1) * nzNA + bA
            For q = rsD0(id) To rsD1(id)
                yy = bY + rdR(q)
                If yy >= 0 And yy < nzGH Then
                    c0 = bX + rdA(q): c1 = bX + rdB(q)
                    If c0 < 0 Then c0 = 0
                    If c1 > nzGW - 1 Then c1 = nzGW - 1
                    For c = c0 To c1
                        occ(s * tamO + yy * nzGW + c) = 1
                    Next
                End If
            Next
            For yy = bY - nzK To bY + rsH(id) + nzK
                If yy >= 0 And yy < nzGH Then
                    o = s * tamP + yy * GW1: acc = 0: ultC = -1
                    For c = 0 To nzGW - 1
                        pre(o + c) = acc
                        If occ(s * tamO + yy * nzGW + c) Then acc = acc + 1: ultC = c
                        ult(s * tamO + yy * nzGW + c) = ultC
                    Next
                    pre(o + nzGW) = acc
                End If
            Next
            nzS(i) = s + 1: nzAI(i) = bA: nzX(i) = bX: nzY(i) = bY
            nzAreaPecas = nzAreaPecas + rsArea(id) * nzRes * nzRes
            If s = nzNS - 1 Then
                If esquerda Then
                    If bX + rsW(id) > altUsada Then altUsada = bX + rsW(id)
                Else
                    If bY + rsH(id) > altUsada Then altUsada = bY + rsH(id)
                End If
            End If
        End If
        If k Mod 10 = 0 Then DoEvents
    Next
End Sub

' imagem da prévia (BMP 24 bits) com todas as chapas lado a lado
Private Sub NzDesenharPrevia()
    Const IW As Long = 760, IH As Long = 520
    Dim img() As Byte, rowB As Long, esc As Double, s As Long, i As Long, q As Long, id As Long
    Dim ox As Long, oy As Long, sw As Long, sh As Long, x As Long, y As Long, c As Long, px As Long, py As Long
    Dim cr As Long, cg As Long, cb As Long, gapPx As Long, f As Integer, arq As String, x0 As Long, x1 As Long, y0 As Long, y1 As Long
    Dim hdr(0 To 53) As Byte
    rowB = ((IW * 3 + 3) \ 4) * 4
    ReDim img(0 To rowB * IH - 1)
    For q = 0 To rowB * IH - 1
        img(q) = 235                                ' fundo cinza claro
    Next
    If nzNS > 0 Then
        gapPx = 10
        esc = (IW - 20 - gapPx * (nzNS - 1)) / (nzNS * (nzSWch / nzRes))
        If (IH - 20) / (nzSHch / nzRes) < esc Then esc = (IH - 20) / (nzSHch / nzRes)
        sw = Int(nzSWch / nzRes * esc): sh = Int(nzSHch / nzRes * esc)
        oy = (IH - sh) \ 2
        For s = 1 To nzNS
            ox = 10 + (s - 1) * (sw + gapPx)
            For y = oy To oy + sh - 1                ' chapa branca com borda
                For x = ox To ox + sw - 1
                    If y = oy Or y = oy + sh - 1 Or x = ox Or x = ox + sw - 1 Then
                        PixelBmp img, rowB, IW, IH, x, y, 120, 120, 120
                    Else
                        PixelBmp img, rowB, IW, IH, x, y, 255, 255, 255
                    End If
                Next
            Next
        Next
        For i = 1 To nzN
            If nzS(i) > 0 Then
                id = (i - 1) * nzNA + nzAI(i)
                CorPeca i, cr, cg, cb
                ox = 10 + (nzS(i) - 1) * (sw + gapPx) + Int(nzMarg / nzRes * esc)
                For q = rsR0(id) To rsR1(id)
                    y0 = oy + Int((nzMarg / nzRes + nzY(i) + ruR(q)) * esc)
                    y1 = oy + Int((nzMarg / nzRes + nzY(i) + ruR(q) + 1) * esc) - 1
                    If y1 < y0 Then y1 = y0
                    x0 = ox + Int((nzX(i) + ruA(q)) * esc)
                    x1 = ox + Int((nzX(i) + ruB(q) + 1) * esc) - 1
                    If x1 < x0 Then x1 = x0
                    For py = y0 To y1
                        For px = x0 To x1
                            PixelBmp img, rowB, IW, IH, px, py, cr, cg, cb
                        Next
                    Next
                Next
            End If
        Next
    End If
    ' cabeçalho BMP
    PoeLong hdr, 0, 0: hdr(0) = 66: hdr(1) = 77
    PoeLong hdr, 2, 54 + rowB * IH
    PoeLong hdr, 10, 54: PoeLong hdr, 14, 40
    PoeLong hdr, 18, IW: PoeLong hdr, 22, IH
    hdr(26) = 1: hdr(28) = 24
    PoeLong hdr, 34, rowB * IH
    arq = Environ$("TEMP") & "\lt_encaixe_" & Format$(Now, "hhnnss") & ".bmp"
    If Len(gPrevArq) > 0 Then
        On Error Resume Next
        Kill gPrevArq
        On Error GoTo 0
    End If
    f = FreeFile
    Open arq For Binary Access Write As #f
    Put #f, , hdr
    Put #f, , img
    Close #f
    gPrevArq = arq
End Sub

Private Sub PixelBmp(img() As Byte, ByVal rowB As Long, ByVal IW As Long, ByVal IH As Long, ByVal x As Long, ByVal y As Long, _
                     ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim o As Long
    If x < 0 Or y < 0 Or x >= IW Or y >= IH Then Exit Sub
    o = y * rowB + x * 3                              ' BMP: linha 0 = de baixo
    img(o) = b: img(o + 1) = g: img(o + 2) = r
End Sub

Private Sub PoeLong(b() As Byte, ByVal pos As Long, ByVal v As Long)
    b(pos) = v And &HFF&: b(pos + 1) = (v \ &H100&) And &HFF&
    b(pos + 2) = (v \ &H10000) And &HFF&: b(pos + 3) = (v \ &H1000000) And &HFF&
End Sub

Private Sub CorPeca(ByVal i As Long, cr As Long, cg As Long, cb As Long)
    Select Case i Mod 8
        Case 0: cr = 52: cg = 101: cb = 164
        Case 1: cr = 204: cg = 102: cb = 0
        Case 2: cr = 78: cg = 154: cb = 6
        Case 3: cr = 117: cg = 80: cb = 123
        Case 4: cr = 193: cg = 125: cb = 17
        Case 5: cr = 32: cg = 150: cb = 160
        Case 6: cr = 164: cg = 0: cb = 0
        Case Else: cr = 85: cg = 87: cb = 83
    End Select
End Sub

' aplica o encaixe calculado no desenho
Private Sub NzAplicar()
    Dim i As Long, id As Long, ox As Double, oy As Double, s As Long, lay As Layer, r As Shape
    Set lay = Camada("CHAPAS")
    For s = 1 To nzNS
        ox = nzSX + nzSW + 50 + (s - 1) * (nzSWch + 50)
        Set r = lay.CreateRectangle2(ox, nzSY, nzSWch, nzSHch)
        r.Fill.ApplyNoFill
        Estilo r, "marca"
    Next
    For i = 1 To nzN
        If nzS(i) > 0 Then
            id = (i - 1) * nzNA + nzAI(i)
            If nzAng(nzAI(i)) <> 0 Then nzShp(i).RotateEx nzAng(nzAI(i)), nzCX(i), nzCY(i)
            ox = nzSX + nzSW + 50 + (nzS(i) - 1) * (nzSWch + 50) + nzMarg
            oy = nzSY + nzMarg
            nzShp(i).Move ox + nzX(i) * nzRes - rsXm(id), oy + nzY(i) * nzRes - rsYm(id)
        End If
    Next
End Sub
