#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste de fumo (smoke test)
==========================

Constrói, em memória, um documento MODELO (com cabeçalho, rodapé, estilo
próprio e página em paisagem) e um documento de ORIGEM, aplica a formatação e
verifica o resultado. Serve para o workflow do GitHub Actions detectar
regressões antes de gerar os executáveis.

Executar localmente:
    python scripts/teste_fumo.py
"""

import os
import sys
import tempfile
import zipfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

from formatar_docx_modelo_gui import COMPONENTES, formatar  # noqa: E402

NS_DOC = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

TIPOS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
{extra}</Types>"""

RELS_RAIZ = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def _estilos(fonte, tamanho_meio_pt, cor_titulo):
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
<w:rPr><w:b/><w:color w:val="%s"/><w:sz w:val="40"/></w:rPr></w:style>
</w:styles>""" % (NS_DOC, fonte, fonte, tamanho_meio_pt, cor_titulo)


def criar_modelo(caminho):
    cabecalho = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr %s><w:p><w:r><w:t>CABEÇALHO DO MODELO §</w:t></w:r></w:p></w:hdr>""" % NS_DOC
    rodape = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr %s><w:p><w:r><w:t>Rodapé do modelo ★</w:t></w:r></w:p></w:ftr>""" % NS_DOC
    documento = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document %s><w:body><w:p><w:r><w:t>modelo</w:t></w:r></w:p>
<w:sectPr>
<w:headerReference w:type="default" r:id="rId10"/>
<w:footerReference w:type="default" r:id="rId11"/>
<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
<w:pgMar w:top="1701" w:right="1701" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr></w:body></w:document>""" % NS_DOC
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>"""
    extra = ('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n'
             '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n')
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", TIPOS.format(extra=extra))
        z.writestr("_rels/.rels", RELS_RAIZ)
        z.writestr("word/document.xml", documento)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/styles.xml", _estilos("Georgia", 26, "0A468C"))
        z.writestr("word/header1.xml", cabecalho)
        z.writestr("word/footer1.xml", rodape)


def criar_origem(caminho):
    documento = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document %s><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Capítulo I</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS" w:hAnsi="Comic Sans MS"/><w:sz w:val="18"/></w:rPr>
<w:t>Conteúdo a preservar: ação, prémio, â.</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Negrito preservado.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="283" w:footer="283" w:gutter="0"/>
</w:sectPr></w:body></w:document>""" % NS_DOC
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", TIPOS.format(extra=""))
        z.writestr("_rels/.rels", RELS_RAIZ)
        z.writestr("word/document.xml", documento)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/styles.xml", _estilos("Arial", 20, "000000"))


def principal():
    with tempfile.TemporaryDirectory() as pasta:
        modelo = os.path.join(pasta, "modelo.docx")
        origem = os.path.join(pasta, "origem.docx")
        saida = os.path.join(pasta, "saida.docx")
        criar_modelo(modelo)
        criar_origem(origem)

        formatar(origem, modelo, saida, set(COMPONENTES), {"fontes"},
                 "1", {}, verboso=False)

        with zipfile.ZipFile(saida) as z:
            nomes = z.namelist()
            documento = z.read("word/document.xml").decode("utf-8")
            estilos = z.read("word/styles.xml").decode("utf-8")
            partes_cab = [n for n in nomes if "header" in n and n.endswith(".xml")]
            cabecalho = z.read(partes_cab[0]).decode("utf-8") if partes_cab else ""

        verificacoes = [
            ("conteúdo preservado", "Conteúdo a preservar: ação, prémio, â." in documento),
            ("negrito preservado", "<w:b/>" in documento),
            ("título preservado", "Capítulo I" in documento),
            ("cabeçalho do modelo importado", "CABEÇALHO DO MODELO" in cabecalho),
            ("referência de cabeçalho reposta", "headerReference" in documento),
            ("referência de rodapé reposta", "footerReference" in documento),
            ("orientação do modelo aplicada", 'w:orient="landscape"' in documento),
            ("margens do modelo aplicadas", 'w:left="1701"' in documento),
            ("tipo de letra do modelo aplicado", "Georgia" in estilos),
            ("formatação directa removida", "Comic Sans MS" not in documento),
        ]

        falhou = False
        for descricao, resultado in verificacoes:
            print("  [%s] %s" % ("OK " if resultado else "FALHA", descricao))
            falhou = falhou or not resultado

        if falhou:
            print("\nTeste de fumo FALHOU.")
            return 1
        print("\nTeste de fumo concluído com sucesso.")
        return 0


if __name__ == "__main__":
    raise SystemExit(principal())
