#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
formatar_docx_modelo_gui.py — versão com interface gráfica
==========================================================

Aplica a formatação de um documento Word MODELO a um ou vários documentos
.docx, preservando integralmente o conteúdo destes.

COMO USAR (VSCode)
------------------
1. Instale a dependência uma única vez, no terminal do VSCode:
       pip install lxml
   (no Windows, se "pip" não for reconhecido: py -m pip install lxml)
2. Abra este ficheiro no VSCode e execute-o (botão ▷ "Run Python File" ou F5).
3. Abre-se uma janela:
       1. escolha o documento MODELO;
       2. adicione um ou mais documentos a formatar (pode seleccionar vários
          de uma vez, com Ctrl ou Shift);
       3. marque o que quer transferir do modelo;
       4. carregue em FORMATAR.
   Os ficheiros resultantes são gravados com o sufixo "_formatado" ao lado dos
   originais, salvo indicação em contrário.

Não é preciso escrever nada na linha de comandos. Ainda assim, quem preferir
pode usá-la: basta executar o ficheiro com argumentos (--help mostra tudo).

O que é transferido (tudo seleccionável na janela)
--------------------------------------------------
  estilos      estilos de parágrafo/caracter/tabela e predefinições de tipo de
               letra; os estilos são associados por NOME, pelo que funciona
               mesmo entre ficheiros com nomes internos diferentes (PT/EN)
  cabecalhos   cabeçalhos do modelo, incluindo logótipos, imagens, símbolos,
               campos e marcas de água, aplicados a todas as secções
  rodapes      idem para os rodapés (numeração de página incluída)
  pagina       dimensão e orientação da página, margens, colunas, bordas de
               página, numeração e "primeira página diferente"
  tema         cores e tipos de letra do tema
  numeracao    listas e marcas do modelo (acrescentadas, sem afectar as da origem)
  definicoes   cabeçalhos pares/ímpares, tabulação por omissão, entre outras

Opcionalmente remove a formatação aplicada à mão nos documentos de origem, para
que os estilos do modelo passem efectivamente a prevalecer. Negrito, itálico e
sublinhado são sempre preservados.

Requisitos: Python 3.8+, lxml e tkinter (incluído no Python do Windows/macOS;
no Linux: sudo apt install python3-tk).
"""


from __future__ import annotations

import argparse
import copy
import json
import os
import posixpath
import shutil
import sys
import zipfile
from collections import OrderedDict

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERRO: é necessária a biblioteca lxml.\n"
        "      Instale com:  pip install lxml\n"
    )
    raise SystemExit(2)


# --------------------------------------------------------------------------- #
# Espaços de nomes
# --------------------------------------------------------------------------- #

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {"w": W, "r": R, "pr": PR, "ct": CT}

RT_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
RT_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
RT_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
RT_FONTTABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
RT_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"

CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_FONTTABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"

CT_POR_EXTENSAO = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "svg": "image/svg+xml",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "webp": "image/webp",
    "ico": "image/x-icon",
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
}

# Ordem imposta pelo esquema (CT_SectPr)
ORDEM_SECTPR = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
    "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
    "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]

# Ordem (parcial mas suficiente) imposta pelo esquema para settings.xml
ORDEM_SETTINGS = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
    "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData",
    "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
    "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
    "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
    "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
    "trackChanges", "doNotTrackMoves", "doNotTrackFormatting",
    "documentProtection", "autoFormatOverride", "styleLockTheme",
    "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
    "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
    "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
    "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
    "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
    "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr",
    "compat", "docVars", "rsids",
]

# Elementos de sectPr copiados pela opção "pagina"
ELEMENTOS_PAGINA = [
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "vAlign", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
]

COMPONENTES = [
    "estilos", "tema", "cabecalhos", "rodapes", "pagina", "numeracao", "definicoes",
]

GRUPOS_LIMPEZA = {
    "fontes": {"rPr": ["rFonts", "sz", "szCs", "position", "kern", "spacing", "w", "effect"]},
    "cores": {"rPr": ["color", "highlight", "shd", "em"]},
    "realce": {"rPr": ["highlight"]},
    "paragrafo": {"pPr": ["spacing", "ind", "jc", "contextualSpacing",
                          "keepNext", "keepLines", "widowControl", "textAlignment"]},
    "bordas": {"pPr": ["pBdr", "shd"]},
    "tabulacoes": {"pPr": ["tabs"]},
}


def q(prefixo: str, nome: str) -> str:
    """Devolve um nome qualificado no formato Clark."""
    return "{%s}%s" % (NS[prefixo], nome)


# --------------------------------------------------------------------------- #
# Pacote OPC (leitura/escrita do ZIP)
# --------------------------------------------------------------------------- #

class Pacote:
    """Representação em memória de um pacote OPC (.docx)."""

    def __init__(self, caminho: str):
        self.caminho = caminho
        self.partes: "OrderedDict[str, bytes]" = OrderedDict()
        with zipfile.ZipFile(caminho, "r") as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                self.partes[info.filename] = z.read(info.filename)

    # -- acesso ---------------------------------------------------------- #

    def existe(self, nome: str) -> bool:
        return nome in self.partes

    def ler(self, nome: str) -> bytes:
        return self.partes[nome]

    def xml(self, nome: str):
        parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
        return etree.fromstring(self.partes[nome], parser)

    def gravar_xml(self, nome: str, elemento) -> None:
        self.partes[nome] = etree.tostring(
            elemento, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def definir(self, nome: str, dados: bytes) -> None:
        self.partes[nome] = dados

    def remover(self, nome: str) -> None:
        self.partes.pop(nome, None)

    def nomes(self):
        return list(self.partes.keys())

    # -- escrita --------------------------------------------------------- #

    def escrever(self, destino: str) -> None:
        ordem = ["[Content_Types].xml"] + [
            n for n in self.partes if n != "[Content_Types].xml"
        ]
        tmp = destino + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for nome in ordem:
                if nome in self.partes:
                    z.writestr(nome, self.partes[nome])
        shutil.move(tmp, destino)


# --------------------------------------------------------------------------- #
# Utilitários de relações e caminhos
# --------------------------------------------------------------------------- #

def caminho_rels(parte: str) -> str:
    d = posixpath.dirname(parte)
    b = posixpath.basename(parte)
    return posixpath.join(d, "_rels", b + ".rels") if d else posixpath.join("_rels", b + ".rels")


def resolver(parte: str, alvo: str) -> str:
    """Resolve um Target relativo à parte que o referencia."""
    base = posixpath.dirname(parte)
    return posixpath.normpath(posixpath.join(base, alvo)).lstrip("/")


def relativizar(parte: str, alvo: str) -> str:
    """Converte um caminho absoluto do pacote em Target relativo à parte."""
    base = posixpath.dirname(parte)
    return posixpath.relpath(alvo, base or ".")


def ler_rels(pacote: Pacote, parte: str):
    nome = caminho_rels(parte)
    if not pacote.existe(nome):
        raiz = etree.Element(q("pr", "Relationships"), nsmap={None: PR})
        return nome, raiz
    return nome, pacote.xml(nome)


def rels_por_id(raiz) -> dict:
    return {rel.get("Id"): rel for rel in raiz.findall(q("pr", "Relationship"))}


def novo_rid(raiz) -> str:
    usados = {rel.get("Id") for rel in raiz.findall(q("pr", "Relationship"))}
    i = 1
    while "rId%d" % i in usados:
        i += 1
    return "rId%d" % i


def rel_por_tipo(raiz, tipo: str):
    for rel in raiz.findall(q("pr", "Relationship")):
        if rel.get("Type") == tipo:
            return rel
    return None


def inserir_ordenado(pai, elemento, ordem) -> None:
    """Insere `elemento` em `pai` respeitando a ordem imposta pelo esquema."""
    nome = etree.QName(elemento).localname
    try:
        pos = ordem.index(nome)
    except ValueError:
        pai.append(elemento)
        return
    for filho in list(pai):
        try:
            p = ordem.index(etree.QName(filho).localname)
        except ValueError:
            continue
        if p > pos:
            filho.addprevious(elemento)
            return
    pai.append(elemento)


def reordenar(pai, ordem) -> None:
    conhecidos, desconhecidos = [], []
    for filho in list(pai):
        nome = etree.QName(filho).localname
        (conhecidos if nome in ordem else desconhecidos).append(filho)
        pai.remove(filho)
    conhecidos.sort(key=lambda e: ordem.index(etree.QName(e).localname))
    for e in conhecidos:
        pai.append(e)
    for e in desconhecidos:
        pai.append(e)


# --------------------------------------------------------------------------- #
# Content types
# --------------------------------------------------------------------------- #

class ContentTypes:
    def __init__(self, pacote: Pacote):
        self.pacote = pacote
        self.raiz = pacote.xml("[Content_Types].xml")

    def defaults(self) -> dict:
        return {
            d.get("Extension", "").lower(): d
            for d in self.raiz.findall(q("ct", "Default"))
        }

    def overrides(self) -> dict:
        return {
            o.get("PartName"): o for o in self.raiz.findall(q("ct", "Override"))
        }

    def garantir_default(self, extensao: str, tipo: str | None = None) -> None:
        ext = extensao.lower().lstrip(".")
        if not ext or ext in self.defaults():
            return
        tipo = tipo or CT_POR_EXTENSAO.get(ext, "application/octet-stream")
        el = etree.SubElement(self.raiz, q("ct", "Default"))
        el.set("Extension", ext)
        el.set("ContentType", tipo)
        # Defaults devem preceder Overrides
        primeiro_override = self.raiz.find(q("ct", "Override"))
        if primeiro_override is not None:
            primeiro_override.addprevious(el)

    def definir_override(self, parte: str, tipo: str) -> None:
        nome = "/" + parte.lstrip("/")
        ovs = self.overrides()
        if nome in ovs:
            ovs[nome].set("ContentType", tipo)
            return
        el = etree.SubElement(self.raiz, q("ct", "Override"))
        el.set("PartName", nome)
        el.set("ContentType", tipo)

    def remover_override(self, parte: str) -> None:
        nome = "/" + parte.lstrip("/")
        ovs = self.overrides()
        if nome in ovs:
            self.raiz.remove(ovs[nome])

    def gravar(self) -> None:
        self.pacote.gravar_xml("[Content_Types].xml", self.raiz)


# --------------------------------------------------------------------------- #
# Inspecção
# --------------------------------------------------------------------------- #

def _twips_para_cm(v) -> str:
    try:
        return "%.2f cm" % (int(v) / 567.0)
    except (TypeError, ValueError):
        return "?"


def inspeccionar(caminho: str) -> None:
    pac = Pacote(caminho)
    print("=" * 72)
    print("Ficheiro: %s" % caminho)
    print("=" * 72)

    doc = pac.xml("word/document.xml")
    _, rels = ler_rels(pac, "word/document.xml")
    mapa = rels_por_id(rels)

    # Estilos
    if pac.existe("word/styles.xml"):
        estilos = pac.xml("word/styles.xml")
        lista = []
        for st in estilos.findall(q("w", "style")):
            nome_el = st.find(q("w", "name"))
            lista.append((
                st.get(q("w", "styleId")),
                nome_el.get(q("w", "val")) if nome_el is not None else "",
                st.get(q("w", "type")) or "paragraph",
            ))
        print("\nESTILOS DEFINIDOS (%d)" % len(lista))
        for sid, nome, tipo in sorted(lista, key=lambda x: (x[2], x[1].lower())):
            print("  [%-9s] %-28s  styleId=%s" % (tipo, nome, sid))
        dd = estilos.find(q("w", "docDefaults"))
        if dd is not None:
            fonte = dd.find(".//" + q("w", "rFonts"))
            tam = dd.find(".//" + q("w", "sz"))
            print("\n  Predefinições: tipo de letra=%s  tamanho=%s pt" % (
                fonte.get(q("w", "ascii")) if fonte is not None else "?",
                (int(tam.get(q("w", "val"))) / 2) if tam is not None else "?",
            ))

    # Secções
    sects = doc.findall(".//" + q("w", "sectPr"))
    print("\nSECÇÕES (%d)" % len(sects))
    for i, sect in enumerate(sects, 1):
        pgsz = sect.find(q("w", "pgSz"))
        pgmar = sect.find(q("w", "pgMar"))
        print("  Secção %d:" % i)
        if pgsz is not None:
            print("    Página: %s x %s  orientação=%s" % (
                _twips_para_cm(pgsz.get(q("w", "w"))),
                _twips_para_cm(pgsz.get(q("w", "h"))),
                pgsz.get(q("w", "orient")) or "portrait",
            ))
        if pgmar is not None:
            print("    Margens: sup=%s inf=%s esq=%s dir=%s cab=%s rod=%s" % (
                _twips_para_cm(pgmar.get(q("w", "top"))),
                _twips_para_cm(pgmar.get(q("w", "bottom"))),
                _twips_para_cm(pgmar.get(q("w", "left"))),
                _twips_para_cm(pgmar.get(q("w", "right"))),
                _twips_para_cm(pgmar.get(q("w", "header"))),
                _twips_para_cm(pgmar.get(q("w", "footer"))),
            ))
        if sect.find(q("w", "titlePg")) is not None:
            print("    1.ª página diferente: sim")
        for ref in sect.findall(q("w", "headerReference")) + sect.findall(q("w", "footerReference")):
            tipo = "cabeçalho" if etree.QName(ref).localname == "headerReference" else "rodapé"
            rel = mapa.get(ref.get(q("r", "id")))
            alvo = resolver("word/document.xml", rel.get("Target")) if rel is not None else "?"
            imgs = _imagens_da_parte(pac, alvo)
            extra = "  [%d imagem(ns): %s]" % (len(imgs), ", ".join(imgs)) if imgs else ""
            print("    %s (%s): %s%s" % (
                tipo, ref.get(q("w", "type")), posixpath.basename(alvo), extra))

    # Numeração e tema
    print("\nOUTRAS PARTES")
    for nome, rotulo in [
        ("word/numbering.xml", "numeração/listas"),
        ("word/theme/theme1.xml", "tema"),
        ("word/fontTable.xml", "tabela de tipos de letra"),
        ("word/settings.xml", "definições"),
    ]:
        print("  %-28s %s" % (rotulo, "presente" if pac.existe(nome) else "ausente"))
    print()


def _imagens_da_parte(pac: Pacote, parte: str):
    nome = caminho_rels(parte)
    if not pac.existe(nome):
        return []
    raiz = pac.xml(nome)
    saida = []
    for rel in raiz.findall(q("pr", "Relationship")):
        if rel.get("TargetMode") == "External":
            continue
        alvo = rel.get("Target", "")
        if "media/" in alvo:
            saida.append(posixpath.basename(alvo))
    return saida


# --------------------------------------------------------------------------- #
# Passo 1 — numeração (fusão aditiva)
# --------------------------------------------------------------------------- #

def fundir_numeracao(destino: Pacote, modelo: Pacote, ct: ContentTypes, log):
    """
    Funde as definições de numeração do modelo nas do destino, deslocando os
    identificadores para não colidirem. Devolve o mapa
    {numId_no_modelo: numId_novo} para remapear referências nas partes
    importadas (estilos, cabeçalhos e rodapés).
    """
    if not modelo.existe("word/numbering.xml"):
        return {}

    num_modelo = modelo.xml("word/numbering.xml")

    if not destino.existe("word/numbering.xml"):
        destino.definir("word/numbering.xml", modelo.ler("word/numbering.xml"))
        ct.definir_override("word/numbering.xml", CT_NUMBERING)
        nome_rels, rels = ler_rels(destino, "word/document.xml")
        if rel_por_tipo(rels, RT_NUMBERING) is None:
            rel = etree.SubElement(rels, q("pr", "Relationship"))
            rel.set("Id", novo_rid(rels))
            rel.set("Type", RT_NUMBERING)
            rel.set("Target", "numbering.xml")
            destino.gravar_xml(nome_rels, rels)
        log("numeração: definições do modelo importadas integralmente")
        return {}

    num_dest = destino.xml("word/numbering.xml")

    def maximo(raiz, tag, attr):
        valores = [0]
        for el in raiz.findall(q("w", tag)):
            try:
                valores.append(int(el.get(q("w", attr))))
            except (TypeError, ValueError):
                pass
        return max(valores)

    desl_abs = maximo(num_dest, "abstractNum", "abstractNumId") + 1
    desl_num = maximo(num_dest, "num", "numId") + 1

    mapa_abs, mapa_num = {}, {}

    abstratos, nums = [], []
    for el in num_modelo.findall(q("w", "abstractNum")):
        antigo = int(el.get(q("w", "abstractNumId")))
        novo = antigo + desl_abs
        mapa_abs[antigo] = novo
        copia = copy.deepcopy(el)
        copia.set(q("w", "abstractNumId"), str(novo))
        # nsid/tmpl duplicados podem confundir o Word: removem-se
        for tag in ("nsid", "tmpl"):
            sub = copia.find(q("w", tag))
            if sub is not None:
                copia.remove(sub)
        abstratos.append(copia)

    for el in num_modelo.findall(q("w", "num")):
        antigo = int(el.get(q("w", "numId")))
        novo = antigo + desl_num
        mapa_num[antigo] = novo
        copia = copy.deepcopy(el)
        copia.set(q("w", "numId"), str(novo))
        ref = copia.find(q("w", "abstractNumId"))
        if ref is not None:
            try:
                ref.set(q("w", "val"), str(mapa_abs[int(ref.get(q("w", "val")))]))
            except (KeyError, TypeError, ValueError):
                pass
        nums.append(copia)

    # abstractNum antes de num (ordem do esquema)
    ultimo_abs = num_dest.findall(q("w", "abstractNum"))
    ancora = ultimo_abs[-1] if ultimo_abs else None
    for el in abstratos:
        if ancora is not None:
            ancora.addnext(el)
            ancora = el
        else:
            num_dest.insert(0, el)
            ancora = el
    for el in nums:
        num_dest.append(el)

    destino.gravar_xml("word/numbering.xml", num_dest)
    if abstratos or nums:
        log("numeração: %d lista(s) do modelo acrescentada(s) sem afectar as existentes"
            % len(nums))
    return mapa_num


def remapear_numeracao(raiz, mapa_num: dict) -> None:
    if not mapa_num:
        return
    for el in raiz.iter(q("w", "numId")):
        try:
            v = int(el.get(q("w", "val")))
        except (TypeError, ValueError):
            continue
        if v in mapa_num:
            el.set(q("w", "val"), str(mapa_num[v]))


# --------------------------------------------------------------------------- #
# Passo 2 — estilos
# --------------------------------------------------------------------------- #

def _nome_estilo(st) -> str:
    el = st.find(q("w", "name"))
    return (el.get(q("w", "val")) if el is not None else "") or ""


def aplicar_estilos(destino: Pacote, modelo: Pacote, ct: ContentTypes,
                    mapa_num: dict, mapa_manual: dict, log):
    """
    Substitui os estilos do destino pelos do modelo, preservando os estilos
    exclusivos do destino, e devolve o mapa {styleId_antigo: styleId_novo}
    para reescrever as referências no corpo do documento.
    """
    if not modelo.existe("word/styles.xml"):
        log("estilos: o modelo não contém styles.xml — passo ignorado")
        return {}

    st_modelo = modelo.xml("word/styles.xml")
    remapear_numeracao(st_modelo, mapa_num)

    estilos_modelo = st_modelo.findall(q("w", "style"))
    por_nome_modelo = {}
    ids_modelo = set()
    for st in estilos_modelo:
        sid = st.get(q("w", "styleId"))
        ids_modelo.add(sid)
        por_nome_modelo[(st.get(q("w", "type")) or "paragraph", _nome_estilo(st).strip().lower())] = sid
        por_nome_modelo[("*", _nome_estilo(st).strip().lower())] = sid

    mapa_ids = {}
    exclusivos = []

    if destino.existe("word/styles.xml"):
        st_dest = destino.xml("word/styles.xml")
        for st in st_dest.findall(q("w", "style")):
            sid = st.get(q("w", "styleId"))
            tipo = st.get(q("w", "type")) or "paragraph"
            nome = _nome_estilo(st).strip()
            chave = nome.lower()

            # Mapeamento manual por nome tem prioridade
            destino_manual = mapa_manual.get(nome) or mapa_manual.get(chave)
            if destino_manual:
                alvo = por_nome_modelo.get((tipo, destino_manual.strip().lower())) \
                    or por_nome_modelo.get(("*", destino_manual.strip().lower())) \
                    or (destino_manual if destino_manual in ids_modelo else None)
                if alvo:
                    mapa_ids[sid] = alvo
                    continue

            alvo = por_nome_modelo.get((tipo, chave))
            if alvo:
                if alvo != sid:
                    mapa_ids[sid] = alvo
                continue

            # Estilo que só existe no documento de origem: preserva-se
            copia = copy.deepcopy(st)
            if sid in ids_modelo:
                novo_id = sid + "Orig"
                n = 1
                while novo_id in ids_modelo:
                    n += 1
                    novo_id = "%sOrig%d" % (sid, n)
                copia.set(q("w", "styleId"), novo_id)
                mapa_ids[sid] = novo_id
                ids_modelo.add(novo_id)
            copia.attrib.pop(q("w", "default"), None)
            exclusivos.append(copia)

    # Reescrever referências internas dos estilos preservados
    for st in exclusivos:
        for tag in ("basedOn", "next", "link"):
            el = st.find(q("w", tag))
            if el is not None:
                v = el.get(q("w", "val"))
                if v in mapa_ids:
                    el.set(q("w", "val"), mapa_ids[v])
        st_modelo.append(st)

    destino.gravar_xml("word/styles.xml", st_modelo)
    ct.definir_override(
        "word/styles.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")

    nome_rels, rels = ler_rels(destino, "word/document.xml")
    if rel_por_tipo(rels, RT_STYLES) is None:
        rel = etree.SubElement(rels, q("pr", "Relationship"))
        rel.set("Id", novo_rid(rels))
        rel.set("Type", RT_STYLES)
        rel.set("Target", "styles.xml")
        destino.gravar_xml(nome_rels, rels)

    log("estilos: %d estilo(s) do modelo aplicados; %d estilo(s) exclusivos da "
        "origem preservados; %d referência(s) remapeada(s)"
        % (len(estilos_modelo), len(exclusivos), len(mapa_ids)))
    return mapa_ids


def aplicar_mapa_estilos(raiz, mapa_ids: dict) -> None:
    if not mapa_ids:
        return
    alvos = (q("w", "pStyle"), q("w", "rStyle"), q("w", "tblStyle"))
    for el in raiz.iter():
        if el.tag in alvos:
            v = el.get(q("w", "val"))
            if v in mapa_ids:
                el.set(q("w", "val"), mapa_ids[v])


# --------------------------------------------------------------------------- #
# Passo 3 — tema e tabela de tipos de letra
# --------------------------------------------------------------------------- #

def aplicar_tema(destino: Pacote, modelo: Pacote, ct: ContentTypes, log):
    copiados = []
    for nome in modelo.nomes():
        if nome.startswith("word/theme/") and nome.endswith(".xml"):
            destino.definir(nome, modelo.ler(nome))
            ct.definir_override(nome, CT_THEME)
            copiados.append(nome)
    if not copiados:
        log("tema: o modelo não contém tema — passo ignorado")
        return

    nome_rels, rels = ler_rels(destino, "word/document.xml")
    if rel_por_tipo(rels, RT_THEME) is None:
        rel = etree.SubElement(rels, q("pr", "Relationship"))
        rel.set("Id", novo_rid(rels))
        rel.set("Type", RT_THEME)
        rel.set("Target", "theme/theme1.xml")
        destino.gravar_xml(nome_rels, rels)

    if modelo.existe("word/fontTable.xml"):
        destino.definir("word/fontTable.xml", modelo.ler("word/fontTable.xml"))
        ct.definir_override("word/fontTable.xml", CT_FONTTABLE)
        nome_rels, rels = ler_rels(destino, "word/document.xml")
        if rel_por_tipo(rels, RT_FONTTABLE) is None:
            rel = etree.SubElement(rels, q("pr", "Relationship"))
            rel.set("Id", novo_rid(rels))
            rel.set("Type", RT_FONTTABLE)
            rel.set("Target", "fontTable.xml")
            destino.gravar_xml(nome_rels, rels)

    log("tema: %s aplicado (cores e tipos de letra do tema)" % ", ".join(
        posixpath.basename(c) for c in copiados))


# --------------------------------------------------------------------------- #
# Passo 4 — cabeçalhos e rodapés (com imagens/logótipos)
# --------------------------------------------------------------------------- #

def _sectpr_modelo(modelo: Pacote, indice):
    doc = modelo.xml("word/document.xml")
    sects = doc.findall(".//" + q("w", "sectPr"))
    if not sects:
        return None
    if isinstance(indice, str) and indice.lower() in ("ultima", "última", "last"):
        return sects[-1]
    try:
        i = int(indice)
    except (TypeError, ValueError):
        i = 1
    i = max(1, min(i, len(sects)))
    return sects[i - 1]


def importar_cabecalhos_rodapes(destino: Pacote, modelo: Pacote, ct: ContentTypes,
                                com_cabecalhos: bool, com_rodapes: bool,
                                sect_modelo, mapa_num: dict, log):
    """
    Remove os cabeçalhos/rodapés do destino e importa os do modelo, incluindo
    imagens, logótipos e demais recursos, repondo as referências em todas as
    secções do documento de destino. Devolve a lista de referências a aplicar.
    """
    if sect_modelo is None:
        log("cabeçalhos/rodapés: modelo sem secções — passo ignorado")
        return [], False, False

    nome_rels_m, rels_m = ler_rels(modelo, "word/document.xml")
    mapa_rels_m = rels_por_id(rels_m)

    # O modelo só substitui o que efectivamente define: se não tiver cabeçalhos
    # (ou rodapés), os do documento de origem são preservados.
    tem_cab = sect_modelo.find(q("w", "headerReference")) is not None
    tem_rod = sect_modelo.find(q("w", "footerReference")) is not None
    if com_cabecalhos and not tem_cab:
        log("cabeçalhos: o modelo não define cabeçalhos — mantidos os da origem")
        com_cabecalhos = False
    if com_rodapes and not tem_rod:
        log("rodapés: o modelo não define rodapés — mantidos os da origem")
        com_rodapes = False
    if not (com_cabecalhos or com_rodapes):
        return [], False, False

    nome_rels_d, rels_d = ler_rels(destino, "word/document.xml")

    # 1) remover cabeçalhos/rodapés existentes no destino
    removidos = 0
    for rel in list(rels_d.findall(q("pr", "Relationship"))):
        tipo = rel.get("Type")
        if tipo == RT_HEADER and not com_cabecalhos:
            continue
        if tipo == RT_FOOTER and not com_rodapes:
            continue
        if tipo in (RT_HEADER, RT_FOOTER):
            alvo = resolver("word/document.xml", rel.get("Target"))
            destino.remover(alvo)
            destino.remover(caminho_rels(alvo))
            ct.remover_override(alvo)
            rels_d.remove(rel)
            removidos += 1

    # 2) importar as partes do modelo referenciadas pela secção escolhida
    referencias = []
    cache_partes = {}
    contador = {"n": 0}

    def importar_parte(parte_modelo: str) -> str:
        """Copia uma parte (cabeçalho/rodapé) e os seus recursos. Devolve o rId."""
        if parte_modelo in cache_partes:
            return cache_partes[parte_modelo]

        contador["n"] += 1
        base = posixpath.basename(parte_modelo)
        novo_nome = "word/mdl_%s" % base
        n = 0
        while destino.existe(novo_nome):
            n += 1
            novo_nome = "word/mdl%d_%s" % (n, base)

        raiz = modelo.xml(parte_modelo)
        remapear_numeracao(raiz, mapa_num)
        destino.gravar_xml(novo_nome, raiz)

        tipo_ct = CT_HEADER if base.startswith("header") else CT_FOOTER
        ct.definir_override(novo_nome, tipo_ct)

        # recursos (imagens, logótipos, hiperligações...)
        nome_r_m = caminho_rels(parte_modelo)
        if modelo.existe(nome_r_m):
            raiz_r = modelo.xml(nome_r_m)
            for rel in raiz_r.findall(q("pr", "Relationship")):
                if rel.get("TargetMode") == "External":
                    continue
                alvo_m = resolver(parte_modelo, rel.get("Target"))
                if not modelo.existe(alvo_m):
                    continue
                base_rec = posixpath.basename(alvo_m)
                dir_rec = posixpath.dirname(alvo_m) or "word"
                novo_rec = posixpath.join(dir_rec, "mdl_%s" % base_rec)
                k = 0
                while destino.existe(novo_rec) and destino.ler(novo_rec) != modelo.ler(alvo_m):
                    k += 1
                    novo_rec = posixpath.join(dir_rec, "mdl%d_%s" % (k, base_rec))
                destino.definir(novo_rec, modelo.ler(alvo_m))
                ext = posixpath.splitext(base_rec)[1].lstrip(".").lower()
                if ext:
                    ct_m = ContentTypes(modelo).defaults().get(ext)
                    ct.garantir_default(
                        ext, ct_m.get("ContentType") if ct_m is not None else None)
                rel.set("Target", relativizar(novo_nome, novo_rec))
            destino.gravar_xml(caminho_rels(novo_nome), raiz_r)

        rid = novo_rid(rels_d)
        rel_nova = etree.SubElement(rels_d, q("pr", "Relationship"))
        rel_nova.set("Id", rid)
        rel_nova.set("Type", RT_HEADER if base.startswith("header") else RT_FOOTER)
        rel_nova.set("Target", relativizar("word/document.xml", novo_nome))
        cache_partes[parte_modelo] = rid
        return rid

    for ref in list(sect_modelo):
        nome_local = etree.QName(ref).localname
        if nome_local == "headerReference" and not com_cabecalhos:
            continue
        if nome_local == "footerReference" and not com_rodapes:
            continue
        if nome_local not in ("headerReference", "footerReference"):
            continue
        rel = mapa_rels_m.get(ref.get(q("r", "id")))
        if rel is None:
            continue
        parte_m = resolver("word/document.xml", rel.get("Target"))
        if not modelo.existe(parte_m):
            continue
        rid = importar_parte(parte_m)
        referencias.append((nome_local, ref.get(q("w", "type")) or "default", rid))

    destino.gravar_xml(nome_rels_d, rels_d)

    if referencias:
        log("cabeçalhos/rodapés: %d removido(s) da origem, %d parte(s) do modelo "
            "importada(s) com os respectivos recursos (imagens/logótipos)"
            % (removidos, len(cache_partes)))
    elif removidos:
        log("cabeçalhos/rodapés: %d removido(s) da origem" % removidos)
    return referencias, com_cabecalhos, com_rodapes


# --------------------------------------------------------------------------- #
# Passo 5 — secções do documento de destino
# --------------------------------------------------------------------------- #

def aplicar_seccoes(doc, sect_modelo, referencias, com_cabecalhos, com_rodapes,
                    com_pagina, log):
    sects = doc.findall(".//" + q("w", "sectPr"))
    if not sects:
        return
    for sect in sects:
        # a) referências de cabeçalho/rodapé
        for el in list(sect):
            nome = etree.QName(el).localname
            if nome == "headerReference" and com_cabecalhos:
                sect.remove(el)
            elif nome == "footerReference" and com_rodapes:
                sect.remove(el)
        for nome_local, tipo, rid in referencias:
            el = etree.Element(q("w", nome_local))
            el.set(q("w", "type"), tipo)
            el.set(q("r", "id"), rid)
            sect.insert(0, el)

        # b) configuração de página
        if com_pagina and sect_modelo is not None:
            for tag in ELEMENTOS_PAGINA:
                existente = sect.find(q("w", tag))
                if existente is not None:
                    sect.remove(existente)
                novo = sect_modelo.find(q("w", tag))
                if novo is not None:
                    sect.append(copy.deepcopy(novo))

        reordenar(sect, ORDEM_SECTPR)

    partes = []
    if referencias:
        partes.append("referências de cabeçalho/rodapé repostas")
    if com_pagina:
        partes.append("configuração de página do modelo aplicada")
    if partes:
        log("secções: %d secção(ões) actualizada(s) — %s" % (len(sects), "; ".join(partes)))


# --------------------------------------------------------------------------- #
# Passo 6 — definições do documento
# --------------------------------------------------------------------------- #

def aplicar_definicoes(destino: Pacote, modelo: Pacote, log):
    if not modelo.existe("word/settings.xml") or not destino.existe("word/settings.xml"):
        return
    s_m = modelo.xml("word/settings.xml")
    s_d = destino.xml("word/settings.xml")
    aplicados = []
    for tag in ("evenAndOddHeaders", "defaultTabStop", "characterSpacingControl",
                "mirrorMargins", "bordersDoNotSurroundHeader",
                "bordersDoNotSurroundFooter", "autoHyphenation", "hyphenationZone"):
        novo = s_m.find(q("w", tag))
        antigo = s_d.find(q("w", tag))
        if novo is None:
            if antigo is not None and tag in ("evenAndOddHeaders", "mirrorMargins"):
                s_d.remove(antigo)
                aplicados.append("-" + tag)
            continue
        if antigo is not None:
            s_d.remove(antigo)
        inserir_ordenado(s_d, copy.deepcopy(novo), ORDEM_SETTINGS)
        aplicados.append(tag)
    destino.gravar_xml("word/settings.xml", s_d)
    if aplicados:
        log("definições: %s" % ", ".join(aplicados))


# --------------------------------------------------------------------------- #
# Passo 7 — limpeza de formatação directa
# --------------------------------------------------------------------------- #

def limpar_formatacao_directa(doc, grupos, log):
    if not grupos:
        return
    alvos_rpr, alvos_ppr = set(), set()
    for g in grupos:
        regra = GRUPOS_LIMPEZA.get(g, {})
        alvos_rpr.update(regra.get("rPr", []))
        alvos_ppr.update(regra.get("pPr", []))

    removidos = 0
    for rpr in doc.iter(q("w", "rPr")):
        for filho in list(rpr):
            if etree.QName(filho).localname in alvos_rpr:
                rpr.remove(filho)
                removidos += 1
    for ppr in doc.iter(q("w", "pPr")):
        for filho in list(ppr):
            if etree.QName(filho).localname in alvos_ppr:
                ppr.remove(filho)
                removidos += 1
    log("limpeza: %d propriedade(s) de formatação directa removida(s) (%s); "
        "negrito, itálico e sublinhado preservados"
        % (removidos, ", ".join(sorted(grupos))))


# --------------------------------------------------------------------------- #
# Orquestração
# --------------------------------------------------------------------------- #

def formatar(origem: str, modelo: str, saida: str, componentes: set,
             limpar: set, seccao_modelo, mapa_manual: dict, verboso: bool = True,
             registo_externo=None):
    registo = []

    def log(msg):
        registo.append(msg)
        if verboso:
            print("  · %s" % msg)
        if registo_externo is not None:
            registo_externo("    %s" % msg)

    pac_m = Pacote(modelo)
    pac_d = Pacote(origem)

    if not pac_d.existe("word/document.xml"):
        raise ValueError("O ficheiro de origem não é um .docx válido (falta word/document.xml).")
    if not pac_m.existe("word/document.xml"):
        raise ValueError("O ficheiro modelo não é um .docx/.dotx válido.")

    ct = ContentTypes(pac_d)

    mapa_num = {}
    if "numeracao" in componentes:
        mapa_num = fundir_numeracao(pac_d, pac_m, ct, log)

    mapa_ids = {}
    if "estilos" in componentes:
        mapa_ids = aplicar_estilos(pac_d, pac_m, ct, mapa_num, mapa_manual, log)

    if "tema" in componentes:
        aplicar_tema(pac_d, pac_m, ct, log)

    sect_m = _sectpr_modelo(pac_m, seccao_modelo)

    referencias = []
    com_cab = "cabecalhos" in componentes
    com_rod = "rodapes" in componentes
    if com_cab or com_rod:
        referencias, com_cab, com_rod = importar_cabecalhos_rodapes(
            pac_d, pac_m, ct, com_cab, com_rod, sect_m, mapa_num, log)

    doc = pac_d.xml("word/document.xml")
    aplicar_mapa_estilos(doc, mapa_ids)
    limpar_formatacao_directa(doc, limpar, log)
    aplicar_seccoes(doc, sect_m, referencias, com_cab, com_rod,
                    "pagina" in componentes, log)
    pac_d.gravar_xml("word/document.xml", doc)

    if "definicoes" in componentes:
        aplicar_definicoes(pac_d, pac_m, log)

    ct.gravar()
    pac_d.escrever(saida)

    # verificação final
    with zipfile.ZipFile(saida) as z:
        etree.fromstring(z.read("word/document.xml"))
    return registo


# --------------------------------------------------------------------------- #
# Interface de linha de comandos
# --------------------------------------------------------------------------- #

def main(argv=None):
    p = argparse.ArgumentParser(
        prog="formatar_docx_por_modelo.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Aplica a formatação de um .docx modelo a outro .docx, "
                    "preservando o conteúdo.",
        epilog="""Exemplos:
  %(prog)s -m modelo.docx -o documento.docx
  %(prog)s -m modelo.docx -o doc.docx -s final.docx --limpar fontes,paragrafo
  %(prog)s -m modelo.docx -o doc.docx --so cabecalhos,rodapes,pagina
  %(prog)s -m modelo.docx -o doc.docx --seccao-modelo ultima
  %(prog)s --inspeccionar modelo.docx
""")
    p.add_argument("--inspeccionar", "-i", metavar="FICHEIRO",
                   help="mostra estilos, secções, cabeçalhos/rodapés e imagens de um .docx")
    p.add_argument("--modelo", "-m", help="ficheiro .docx/.dotx cuja formatação se quer replicar")
    p.add_argument("--origem", "-o", help="ficheiro .docx a formatar (conteúdo preservado)")
    p.add_argument("--saida", "-s", help="ficheiro a produzir (por omissão: <origem>_formatado.docx)")
    p.add_argument("--so", metavar="LISTA",
                   help="aplicar APENAS estes componentes (separados por vírgula): "
                        + ", ".join(COMPONENTES))
    for c in COMPONENTES:
        p.add_argument("--sem-%s" % c, action="store_true",
                       help="não aplicar o componente '%s'" % c)
    p.add_argument("--limpar", metavar="LISTA", default="",
                   help="remover formatação directa do corpo: "
                        + ", ".join(sorted(GRUPOS_LIMPEZA)) + ", tudo")
    p.add_argument("--seccao-modelo", default="1", metavar="N",
                   help="secção do modelo de onde se extraem cabeçalhos/rodapés e "
                        "configuração de página (N ou 'ultima'; por omissão 1)")
    p.add_argument("--mapa-estilos", metavar="FICHEIRO.json",
                   help='mapeamento manual de nomes de estilo, p. ex. '
                        '{"Corpo": "Corpo de texto"}')
    p.add_argument("--silencioso", action="store_true", help="não escrever o relatório")

    a = p.parse_args(argv)

    if a.inspeccionar:
        inspeccionar(a.inspeccionar)
        return 0

    if not a.modelo or not a.origem:
        p.error("são obrigatórios --modelo e --origem (ou use --inspeccionar).")
    for f in (a.modelo, a.origem):
        if not os.path.isfile(f):
            p.error("ficheiro inexistente: %s" % f)

    saida = a.saida or (os.path.splitext(a.origem)[0] + "_formatado.docx")
    if os.path.abspath(saida) in (os.path.abspath(a.origem), os.path.abspath(a.modelo)):
        p.error("o ficheiro de saída não pode coincidir com a origem nem com o modelo.")

    if a.so:
        componentes = {c.strip().lower() for c in a.so.split(",") if c.strip()}
        desconhecidos = componentes - set(COMPONENTES)
        if desconhecidos:
            p.error("componente(s) desconhecido(s): %s" % ", ".join(sorted(desconhecidos)))
    else:
        componentes = set(COMPONENTES)
        for c in COMPONENTES:
            if getattr(a, "sem_%s" % c):
                componentes.discard(c)
    if not componentes:
        p.error("não foi seleccionado nenhum componente de formatação.")

    limpar = {g.strip().lower() for g in a.limpar.split(",") if g.strip()}
    if "tudo" in limpar:
        limpar = set(GRUPOS_LIMPEZA)
    desconhecidos = limpar - set(GRUPOS_LIMPEZA)
    if desconhecidos:
        p.error("grupo(s) de limpeza desconhecido(s): %s" % ", ".join(sorted(desconhecidos)))

    mapa_manual = {}
    if a.mapa_estilos:
        with open(a.mapa_estilos, "r", encoding="utf-8") as fh:
            mapa_manual = json.load(fh)

    if not a.silencioso:
        print("Modelo : %s" % a.modelo)
        print("Origem : %s" % a.origem)
        print("Saída  : %s" % saida)
        print("Componentes: %s" % ", ".join(sorted(componentes)))
        print("-" * 72)

    try:
        formatar(a.origem, a.modelo, saida, componentes, limpar,
                 a.seccao_modelo, mapa_manual, verboso=not a.silencioso)
    except Exception as exc:  # noqa: BLE001
        sys.stderr.write("ERRO: %s\n" % exc)
        return 1

    if not a.silencioso:
        print("-" * 72)
        print("Concluído: %s" % saida)
    return 0


def processar_lote(modelo, ficheiros, pasta_saida, sufixo, componentes, limpar,
                   seccao_modelo, mapa_manual, log, substituir=False):
    """Aplica o modelo a vários ficheiros. Devolve (sucessos, falhas)."""
    sucessos, falhas = [], []
    for i, origem in enumerate(ficheiros, 1):
        nome = os.path.basename(origem)
        log("[%d/%d] %s" % (i, len(ficheiros), nome))
        try:
            base, ext = os.path.splitext(nome)
            if substituir:
                destino = os.path.join(os.path.dirname(origem) or ".",
                                       base + "__tmp_formatado" + ext)
            else:
                pasta = pasta_saida or os.path.dirname(origem) or "."
                os.makedirs(pasta, exist_ok=True)
                destino = os.path.join(pasta, base + sufixo + ext)
                if os.path.abspath(destino) == os.path.abspath(origem):
                    destino = os.path.join(pasta, base + "_formatado" + ext)

            formatar(origem, modelo, destino, componentes, limpar,
                     seccao_modelo, mapa_manual,
                     verboso=False, registo_externo=log)

            if substituir:
                backup = origem + ".bak"
                if os.path.exists(backup):
                    os.remove(backup)
                shutil.copy2(origem, backup)
                shutil.move(destino, origem)
                destino = origem
                log("    cópia de segurança: %s" % os.path.basename(backup))
            log("    → %s" % os.path.basename(destino))
            sucessos.append(destino)
        except Exception as exc:  # noqa: BLE001
            log("    ERRO: %s" % exc)
            falhas.append((origem, str(exc)))
    return sucessos, falhas


# --------------------------------------------------------------------------- #
# Interface gráfica (tkinter)
# --------------------------------------------------------------------------- #

ROTULOS_COMPONENTES = [
    ("estilos", "Estilos e tipos de letra (parágrafos, títulos, tabelas)"),
    ("cabecalhos", "Cabeçalhos — com logótipos, imagens e símbolos"),
    ("rodapes", "Rodapés — com numeração de página e símbolos"),
    ("pagina", "Configuração de página (margens, orientação, colunas)"),
    ("tema", "Tema (cores e tipos de letra do tema)"),
    ("numeracao", "Numeração e marcas de lista"),
    ("definicoes", "Definições do documento (pares/ímpares, tabulações)"),
]

ROTULOS_LIMPEZA = [
    ("fontes", "Tipos de letra e tamanhos aplicados à mão"),
    ("cores", "Cores de texto e sombreados"),
    ("paragrafo", "Alinhamentos, avanços e espaçamentos"),
    ("bordas", "Limites e sombreados de parágrafo"),
    ("tabulacoes", "Tabulações manuais"),
]


def iniciar_interface():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
    except ImportError:
        sys.stderr.write(
            "ERRO: a interface gráfica precisa do módulo tkinter.\n"
            "      Windows/macOS: reinstale o Python com a opção 'tcl/tk'.\n"
            "      Linux: sudo apt install python3-tk\n"
            "      Em alternativa, use a linha de comandos (--help).\n")
        return 2

    import queue
    import threading

    class Aplicacao:
        def __init__(self, raiz):
            self.raiz = raiz
            self.fila = queue.Queue()
            self.modelo = tk.StringVar()
            self.pasta_saida = tk.StringVar()
            self.sufixo = tk.StringVar(value="_formatado")
            self.substituir = tk.BooleanVar(value=False)
            self.seccao = tk.StringVar(value="1")
            self.ultima_pasta = os.path.expanduser("~")
            self.componentes = {c: tk.BooleanVar(value=True) for c, _ in ROTULOS_COMPONENTES}
            self.limpeza = {g: tk.BooleanVar(value=False) for g, _ in ROTULOS_LIMPEZA}
            self.a_trabalhar = False

            raiz.title("Formatar documentos Word a partir de um modelo")
            raiz.minsize(820, 660)
            try:
                raiz.call("ttk::style", "theme", "use", "clam")
            except tk.TclError:
                pass

            self._construir()
            self.raiz.after(120, self._drenar_fila)

        # -- construção da janela ------------------------------------- #

        def _construir(self):
            pad = {"padx": 10, "pady": 6}
            topo = ttk.Frame(self.raiz, padding=12)
            topo.pack(fill="both", expand=True)

            # 1. modelo
            cx1 = ttk.LabelFrame(topo, text=" 1. Documento MODELO (a formatação a copiar) ")
            cx1.pack(fill="x", **pad)
            lin = ttk.Frame(cx1); lin.pack(fill="x", padx=10, pady=10)
            ttk.Entry(lin, textvariable=self.modelo).pack(side="left", fill="x", expand=True)
            ttk.Button(lin, text="Escolher…", command=self.escolher_modelo).pack(side="left", padx=(8, 0))
            ttk.Button(lin, text="Ver formatações", command=self.inspeccionar_modelo).pack(side="left", padx=(6, 0))

            # 2. documentos
            cx2 = ttk.LabelFrame(topo, text=" 2. Documentos A FORMATAR (o conteúdo é preservado) ")
            cx2.pack(fill="both", expand=True, **pad)
            corpo = ttk.Frame(cx2); corpo.pack(fill="both", expand=True, padx=10, pady=10)
            barra = ttk.Scrollbar(corpo, orient="vertical")
            self.lista = tk.Listbox(corpo, selectmode="extended", height=6,
                                    yscrollcommand=barra.set, activestyle="none")
            barra.config(command=self.lista.yview)
            self.lista.pack(side="left", fill="both", expand=True)
            barra.pack(side="left", fill="y")
            botoes = ttk.Frame(corpo); botoes.pack(side="left", fill="y", padx=(8, 0))
            ttk.Button(botoes, text="Adicionar…", command=self.adicionar).pack(fill="x", pady=2)
            ttk.Button(botoes, text="Remover", command=self.remover).pack(fill="x", pady=2)
            ttk.Button(botoes, text="Limpar", command=self.limpar_lista).pack(fill="x", pady=2)

            # 3. opções
            cx3 = ttk.LabelFrame(topo, text=" 3. O que transferir do modelo ")
            cx3.pack(fill="x", **pad)
            grelha = ttk.Frame(cx3); grelha.pack(fill="x", padx=10, pady=(10, 4))
            for i, (chave, rotulo) in enumerate(ROTULOS_COMPONENTES):
                ttk.Checkbutton(grelha, text=rotulo, variable=self.componentes[chave]).grid(
                    row=i % 4, column=i // 4, sticky="w", padx=(0, 18), pady=1)

            lin = ttk.Frame(cx3); lin.pack(fill="x", padx=10, pady=(2, 10))
            ttk.Label(lin, text="Secção do modelo a usar para cabeçalhos/página:").pack(side="left")
            ttk.Combobox(lin, textvariable=self.seccao, width=10, state="readonly",
                         values=["1", "2", "3", "ultima"]).pack(side="left", padx=6)

            cx4 = ttk.LabelFrame(
                topo, text=" 4. Remover formatação aplicada à mão nos documentos "
                           "(opcional — negrito, itálico e sublinhado são sempre preservados) ")
            cx4.pack(fill="x", **pad)
            grelha = ttk.Frame(cx4); grelha.pack(fill="x", padx=10, pady=10)
            for i, (chave, rotulo) in enumerate(ROTULOS_LIMPEZA):
                ttk.Checkbutton(grelha, text=rotulo, variable=self.limpeza[chave]).grid(
                    row=i % 3, column=i // 3, sticky="w", padx=(0, 18), pady=1)

            # 5. destino
            cx5 = ttk.LabelFrame(topo, text=" 5. Onde guardar ")
            cx5.pack(fill="x", **pad)
            lin = ttk.Frame(cx5); lin.pack(fill="x", padx=10, pady=(10, 2))
            ttk.Label(lin, text="Pasta:").pack(side="left")
            ttk.Entry(lin, textvariable=self.pasta_saida).pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(lin, text="Escolher…", command=self.escolher_pasta).pack(side="left")
            lin = ttk.Frame(cx5); lin.pack(fill="x", padx=10, pady=(2, 10))
            ttk.Label(lin, text="(vazio = ao lado do original)   Sufixo:").pack(side="left")
            ttk.Entry(lin, textvariable=self.sufixo, width=16).pack(side="left", padx=6)
            ttk.Checkbutton(lin, text="Substituir os ficheiros originais (guarda cópia .bak)",
                            variable=self.substituir, command=self._alternar_substituir
                            ).pack(side="left", padx=12)

            # acção + registo
            lin = ttk.Frame(topo); lin.pack(fill="x", **pad)
            self.botao = ttk.Button(lin, text="FORMATAR", command=self.executar)
            self.botao.pack(side="left")
            ttk.Button(lin, text="Abrir pasta de destino", command=self.abrir_pasta).pack(side="left", padx=8)
            ttk.Button(lin, text="Sair", command=self.raiz.destroy).pack(side="right")

            cx6 = ttk.LabelFrame(topo, text=" Registo ")
            cx6.pack(fill="both", expand=True, **pad)
            quadro = ttk.Frame(cx6); quadro.pack(fill="both", expand=True, padx=10, pady=10)
            barra = ttk.Scrollbar(quadro, orient="vertical")
            self.registo = tk.Text(quadro, height=10, wrap="word", yscrollcommand=barra.set)
            barra.config(command=self.registo.yview)
            self.registo.pack(side="left", fill="both", expand=True)
            barra.pack(side="left", fill="y")
            self.escrever("Escolha o documento modelo e os documentos a formatar.")

        # -- auxiliares ------------------------------------------------ #

        def escrever(self, texto):
            self.registo.insert("end", texto + "\n")
            self.registo.see("end")

        def _drenar_fila(self):
            try:
                while True:
                    tipo, dados = self.fila.get_nowait()
                    if tipo == "log":
                        self.escrever(dados)
                    elif tipo == "fim":
                        self.a_trabalhar = False
                        self.botao.config(state="normal", text="FORMATAR")
                        sucessos, falhas = dados
                        self.escrever("-" * 60)
                        self.escrever("Concluído: %d ficheiro(s) formatado(s), %d erro(s)."
                                      % (len(sucessos), len(falhas)))
                        from tkinter import messagebox as mb
                        if falhas:
                            mb.showwarning("Concluído com erros",
                                           "%d ficheiro(s) formatado(s).\n%d erro(s) — "
                                           "ver registo." % (len(sucessos), len(falhas)))
                        else:
                            mb.showinfo("Concluído",
                                        "%d ficheiro(s) formatado(s) com sucesso."
                                        % len(sucessos))
            except queue.Empty:
                pass
            self.raiz.after(120, self._drenar_fila)

        def _alternar_substituir(self):
            estado = "disabled" if self.substituir.get() else "normal"
            del estado  # a entrada permanece activa; a opção apenas prevalece

        def escolher_modelo(self):
            f = filedialog.askopenfilename(
                title="Escolher o documento modelo",
                initialdir=self.ultima_pasta,
                filetypes=[("Documentos Word", "*.docx *.dotx"), ("Todos os ficheiros", "*.*")])
            if f:
                self.modelo.set(f)
                self.ultima_pasta = os.path.dirname(f)
                self.escrever("Modelo: %s" % f)

        def adicionar(self):
            fs = filedialog.askopenfilenames(
                title="Escolher os documentos a formatar (pode seleccionar vários)",
                initialdir=self.ultima_pasta,
                filetypes=[("Documentos Word", "*.docx"), ("Todos os ficheiros", "*.*")])
            existentes = set(self.lista.get(0, "end"))
            for f in fs:
                if f not in existentes:
                    self.lista.insert("end", f)
                self.ultima_pasta = os.path.dirname(f)

        def remover(self):
            for i in reversed(self.lista.curselection()):
                self.lista.delete(i)

        def limpar_lista(self):
            self.lista.delete(0, "end")

        def escolher_pasta(self):
            d = filedialog.askdirectory(title="Pasta de destino", initialdir=self.ultima_pasta)
            if d:
                self.pasta_saida.set(d)

        def abrir_pasta(self):
            pasta = self.pasta_saida.get().strip()
            if not pasta:
                itens = self.lista.get(0, "end")
                pasta = os.path.dirname(itens[0]) if itens else ""
            if not pasta or not os.path.isdir(pasta):
                messagebox.showinfo("Pasta", "Ainda não há pasta de destino definida.")
                return
            try:
                if sys.platform.startswith("win"):
                    os.startfile(pasta)  # noqa: S606
                elif sys.platform == "darwin":
                    os.system('open "%s"' % pasta)
                else:
                    os.system('xdg-open "%s"' % pasta)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Pasta", str(exc))

        def inspeccionar_modelo(self):
            modelo = self.modelo.get().strip()
            if not os.path.isfile(modelo):
                messagebox.showwarning("Modelo", "Escolha primeiro um documento modelo.")
                return
            import io
            import contextlib
            buffer = io.StringIO()
            try:
                with contextlib.redirect_stdout(buffer):
                    inspeccionar(modelo)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Modelo", "Não foi possível ler o modelo:\n%s" % exc)
                return
            janela = tk.Toplevel(self.raiz)
            janela.title("Formatações disponíveis no modelo")
            janela.geometry("760x560")
            barra = ttk.Scrollbar(janela, orient="vertical")
            texto = tk.Text(janela, wrap="none", yscrollcommand=barra.set)
            barra.config(command=texto.yview)
            texto.pack(side="left", fill="both", expand=True)
            barra.pack(side="left", fill="y")
            texto.insert("1.0", buffer.getvalue())
            texto.config(state="disabled")

        # -- execução -------------------------------------------------- #

        def executar(self):
            if self.a_trabalhar:
                return
            modelo = self.modelo.get().strip()
            ficheiros = list(self.lista.get(0, "end"))
            if not os.path.isfile(modelo):
                messagebox.showwarning("Modelo", "Escolha o documento modelo.")
                return
            if not ficheiros:
                messagebox.showwarning("Documentos", "Adicione pelo menos um documento a formatar.")
                return
            if any(os.path.abspath(f) == os.path.abspath(modelo) for f in ficheiros):
                messagebox.showwarning(
                    "Documentos", "O modelo não pode constar da lista de documentos a formatar.")
                return
            antigos = [f for f in ficheiros if f.lower().endswith(".doc")]
            if antigos:
                messagebox.showwarning(
                    "Formato antigo",
                    "Os ficheiros .doc não são suportados. Abra-os no Word e "
                    "guarde-os como .docx:\n\n" + "\n".join(os.path.basename(f) for f in antigos))
                return
            componentes = {c for c, v in self.componentes.items() if v.get()}
            if not componentes:
                messagebox.showwarning("Opções", "Seleccione pelo menos uma formatação a transferir.")
                return
            if self.substituir.get() and not messagebox.askyesno(
                    "Substituir originais",
                    "Os ficheiros originais vão ser substituídos.\n"
                    "É guardada uma cópia de segurança com a extensão .bak.\n\nContinuar?"):
                return

            limpar = {g for g, v in self.limpeza.items() if v.get()}
            self.a_trabalhar = True
            self.botao.config(state="disabled", text="a formatar…")
            self.registo.delete("1.0", "end")
            self.escrever("Modelo: %s" % os.path.basename(modelo))
            self.escrever("Formatações: %s" % ", ".join(sorted(componentes)))
            if limpar:
                self.escrever("Limpeza de formatação directa: %s" % ", ".join(sorted(limpar)))
            self.escrever("-" * 60)

            def tarefa():
                def log(msg):
                    self.fila.put(("log", msg))
                try:
                    resultado = processar_lote(
                        modelo, ficheiros, self.pasta_saida.get().strip(),
                        self.sufixo.get().strip() or "_formatado",
                        componentes, limpar, self.seccao.get(), {}, log,
                        substituir=self.substituir.get())
                except Exception as exc:  # noqa: BLE001
                    log("ERRO GERAL: %s" % exc)
                    resultado = ([], [("", str(exc))])
                self.fila.put(("fim", resultado))

            threading.Thread(target=tarefa, daemon=True).start()

    raiz = tk.Tk()
    Aplicacao(raiz)
    raiz.mainloop()
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        raise SystemExit(main())
    raise SystemExit(iniciar_interface())

