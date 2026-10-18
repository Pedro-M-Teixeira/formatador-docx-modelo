# Formatador DOCX por Modelo
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.23457134.svg)](https://doi.org/10.5281/zenodo.23457134)
[![Licença: MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-blue.svg)](LICENSE)

Cita este software: 
Marques Teixeira, P. (2026). Formatador documentos docx (Version v1.0.2) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.23457134


Aplica a formatação de um documento Word **modelo** a um ou vários documentos
`.docx`, **sem alterar o conteúdo destes**. Pensado para quem tem de normalizar
dezenas de documentos segundo um modelo institucional: estilos, cabeçalhos com
logótipos, rodapés com numeração de página, margens e orientação passam a ser
iguais aos do modelo; o texto, as tabelas, as imagens do corpo, os campos e as
notas ficam exactamente como estavam.

> Substitua `UTILIZADOR/REPOSITORIO` no endereço do emblema acima pelo nome do
> seu utilizador e do seu repositório no GitHub.

![Janela da aplicação](assets/icone.png)

---

## O que é transferido do modelo

Tudo seleccionável na janela, caso a caso:

| Componente | O que inclui |
|---|---|
| **Estilos e tipos de letra** | estilos de parágrafo, caracter e tabela e predefinições de tipo de letra e corpo. A associação é feita por **nome**, pelo que funciona mesmo entre ficheiros com nomes internos diferentes (p. ex. `Ttulo1` ↔ `Heading1`). Os estilos que só existam no documento de origem são preservados. |
| **Cabeçalhos** | o cabeçalho do modelo com logótipos, imagens, símbolos, campos e caixas de texto, aplicado a todas as secções, distinguindo primeira página e páginas pares/ímpares. |
| **Rodapés** | idem, numeração de página incluída. |
| **Configuração de página** | dimensão e orientação, margens (incluindo as de cabeçalho e rodapé), colunas, limites de página, numeração e «primeira página diferente». |
| **Tema** | cores e tipos de letra do tema. |
| **Numeração e marcas** | listas do modelo, acrescentadas sem afectar as numerações já existentes no documento. |
| **Definições do documento** | cabeçalhos pares/ímpares, tabulação por omissão, entre outras. |

Opcionalmente, remove a formatação aplicada à mão nos documentos de origem — é
o que resolve o caso frequente de «os estilos do modelo não pegam», já que o
Word dá prioridade à formatação directa sobre os estilos. O negrito, o itálico
e o sublinhado nunca são removidos.

---

## Utilizar sem instalar o Python

1. Abra a secção **[https://github.com/Pedro-M-Teixeira/formatador-docx-modelo/releases/latest)]** deste repositório.
2. Descarregue o ficheiro correspondente ao seu sistema:
   - Windows → `FormatadorDocx-Windows.zip`
   - macOS → `FormatadorDocx-macOS.zip`
   - Linux → `FormatadorDocx-Linux.tar.gz`
3. Extraia o ficheiro e execute `FormatadorDocx`.

**Windows** — na primeira execução o SmartScreen pode avisar que o editor é
desconhecido, por o executável não estar assinado digitalmente: escolha
«Mais informações» → «Executar mesmo assim». Alguns antivírus assinalam
executáveis gerados com PyInstaller como suspeitos; é um falso positivo
conhecido, e quem quiser confirmar pode compilar o executável a partir do
código-fonte com as instruções abaixo.

**macOS** — na primeira execução, clique no ficheiro com o botão direito do
rato e escolha «Abrir», para contornar o Gatekeeper.

### Como usar a janela

1. **Documento MODELO** → «Escolher…». O botão «Ver formatações» mostra o que
   esse modelo contém: estilos, margens, orientação, cabeçalhos, rodapés e as
   imagens que estes incluem.
2. **Documentos a formatar** → «Adicionar…». Pode seleccionar vários de uma vez
   (Ctrl ou Shift) e repetir a operação noutras pastas.
3. Marque o que quer transferir do modelo.
4. Em **Onde guardar**, deixe a pasta vazia para gravar ao lado dos originais
   com o sufixo `_formatado`. A opção de substituir os originais guarda sempre
   uma cópia de segurança `.bak`.
5. **FORMATAR**. O registo indica, documento a documento, o que foi aplicado.

Na pasta `exemplos/` há um modelo e um documento para experimentar.

---

## Utilizar com o Python instalado

```bash
pip install lxml
python formatar_docx_modelo_gui.py          # abre a janela
```

No Linux é ainda necessário o tkinter: `sudo apt install python3-tk`.

### Linha de comandos

O mesmo ficheiro aceita argumentos; e `formatar_docx_por_modelo.py` é a versão
exclusivamente de linha de comandos, útil em automatismos:

```bash
python formatar_docx_por_modelo.py --inspeccionar modelo.docx
python formatar_docx_por_modelo.py -m modelo.docx -o documento.docx
python formatar_docx_por_modelo.py -m modelo.docx -o doc.docx -s final.docx \
    --limpar fontes,paragrafo
python formatar_docx_por_modelo.py -m modelo.docx -o doc.docx \
    --so cabecalhos,rodapes,pagina
python formatar_docx_por_modelo.py --help
```

---

## Publicar este projecto no GitHub

Com o [Git](https://git-scm.com/downloads) instalado, dentro da pasta do
projecto:

```bash
git init
git add .
git commit -m "Versão inicial"
git branch -M main
git remote add origin https://github.com/UTILIZADOR/REPOSITORIO.git
git push -u origin main
```

A partir daqui, **cada `push` compila automaticamente** os executáveis para
Windows, macOS e Linux. Pode descarregá-los no separador **Actions** →
último trabalho → secção *Artifacts* (úteis para testar).

Para publicar uma versão destinada a terceiros, crie uma etiqueta:

```bash
git tag v1.0.0
git push origin v1.0.0
```

O GitHub Actions compila os três executáveis e cria sozinho uma **Release**
com os ficheiros prontos a distribuir. É esse o endereço a enviar a quem não
tem Python instalado. Nas versões seguintes, actualize os números em
`version_info.txt` e em `CHANGELOG.md` antes de criar a nova etiqueta.

> Se a Release não chegar a ser criada, confirme em *Settings → Actions →
> General → Workflow permissions* que está activa a opção
> *Read and write permissions*.

---

## Compilar o executável no seu computador

```bash
# Windows: duplo clique em scripts\compilar_windows.bat
# macOS/Linux:
bash scripts/compilar_unix.sh
```

O executável fica em `dist/`. Em alternativa, manualmente:

```bash
pip install -r requirements.txt pyinstaller
pyinstaller build.spec --noconfirm --clean
```

---

## Estrutura do repositório

```
.
├── formatar_docx_modelo_gui.py   # aplicação principal (janela + linha de comandos)
├── formatar_docx_por_modelo.py   # versão só de linha de comandos
├── build.spec                    # especificação do PyInstaller
├── version_info.txt              # metadados do executável Windows
├── requirements.txt              # dependências (lxml)
├── assets/                       # ícone da aplicação
├── exemplos/                     # modelo e documento de demonstração
├── scripts/
│   ├── compilar_windows.bat      # compilação local no Windows
│   ├── compilar_unix.sh          # compilação local em macOS/Linux
│   └── teste_fumo.py             # teste automático usado na compilação
└── .github/workflows/build.yml   # compilação e publicação automáticas
```

---

## Como funciona

O programa não converte nem reescreve os documentos: manipula directamente as
partes XML do pacote OOXML (`word/document.xml`, `word/styles.xml`,
`word/header*.xml`, relações e tipos de conteúdo), substituindo apenas o que
diz respeito a formatação. As imagens dos cabeçalhos e rodapés do modelo são
copiadas com as respectivas relações, pelo que logótipos e marcas de água
seguem intactos.

### Limitações conhecidas

- Só trabalha com `.docx` e `.dotx`. Ficheiros `.doc` antigos têm de ser
  abertos no Word e guardados como `.docx`.
- Todas as secções do documento recebem os cabeçalhos e rodapés da secção do
  modelo escolhida na janela; não é feita correspondência secção a secção.
- Documentos protegidos ou com palavra-passe não são processados.
- Se o modelo não definir cabeçalhos ou rodapés, os do documento de origem são
  mantidos.

Recomenda-se, como em qualquer processamento em lote, guardar os resultados em
pasta separada antes de substituir os originais.

---

## Licença

Distribuído sob a licença [MIT](LICENSE).
