# Registo de alterações

Todas as alterações assinaláveis deste projecto são documentadas neste ficheiro.
O formato segue o [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/) e
a numeração segue o [Versionamento Semântico](https://semver.org/lang/pt-BR/).

## [1.0.0] — 2026-10-13

### Adicionado
- Interface gráfica para escolher o documento modelo e vários documentos a formatar.
- Transferência selectiva de estilos, cabeçalhos, rodapés, configuração de
  página, tema, numeração e definições do documento.
- Importação de cabeçalhos e rodapés com logótipos, imagens, símbolos e campos.
- Associação de estilos por nome, compatível com ficheiros em idiomas diferentes.
- Remoção opcional de formatação aplicada à mão, preservando negrito, itálico
  e sublinhado.
- Substituição dos ficheiros originais com cópia de segurança `.bak`.
- Versão para linha de comandos (`formatar_docx_por_modelo.py`).
- Compilação automática de executáveis para Windows, macOS e Linux.
