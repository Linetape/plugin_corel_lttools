# LT Tools — plugin para CorelDRAW

Barra de ferramentas VBA para CorelDRAW 2026 (Graphics Suite e Technical Suite), da **Line Tape Iluminação e Locação**, inspirada no eCut: medição e orçamento, encaixe de peças (retangular e pela forma, com giro e prévia), corte em plotter/laser/CNC, letra caixa/LED/neon, geradores e edição.

## Instalar

1. Feche **todos** os CorelDRAW abertos.
2. Rode `instalador/LTTools_Instalador.exe`.
3. Abra o CorelDRAW (só uma cópia). A barra **LT Tools** aparece; arraste-a para o topo.

Se a barra não aparecer: abra uma vez *Ferramentas > Scripts > Scripts* e desligue *Ferramentas > Opções > CorelDRAW > VBA > Retardar carregamento de VBA*.

## Atualizar

Clique no botão **Atualizar** (ícone verde) da barra. Ele lê `versao.txt` deste repositório; se houver versão nova, baixa os arquivos de `arquivos.txt` para `Documentos\Corel\Corel Content\LTTools` e troca o código no CorelDRAW.

> O download direto só funciona com o repositório **público** (o plugin não guarda senha/token). Com o repositório privado, o botão usa os arquivos que estiverem na pasta local.

## Publicar uma versão nova

1. Altere `LTTools.bas` / `frmLT_codigo.txt` (arquivos em **Windows-1252, CRLF** — não converter para UTF-8).
2. Suba a constante `LT_VERSAO` em `LTTools.bas` e o número em `versao.txt` (os dois iguais).
3. Ícones novos: gere com `ferramentas/icones.py`, coloque em `icones/` e inclua em `arquivos.txt`.
4. Faça o commit/push na `main`. Os CorelDRAW instalados recebem ao clicar em **Atualizar**.
5. (Opcional) Gere o instalador novo: atualize o Corel, feche-o e rode `instalador/build.bat`.

## Estrutura

| Caminho | O que é |
|---|---|
| `LTTools.bas` | módulo principal (todas as ferramentas) |
| `frmLT_codigo.txt` | código do formulário `frmLT` (janelas criadas em tempo de execução) |
| `LTUpdate.bas` | atualizador (botão Atualizar) |
| `icones/` | ícones da barra (`.png`) e dos menus (`.bmp`/`.ico`) |
| `versao.txt`, `arquivos.txt` | usados pelo atualizador |
| `instalador/` | instalador `.exe`, código C# e `build.bat` |
| `ferramentas/` | scripts auxiliares (gerador de ícones, protótipo do encaixe) |
