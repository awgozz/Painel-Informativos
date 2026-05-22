# Gerador de Informativos

App em Python para preencher informativos da transportadora, copiar o texto completo para o WhatsApp e gerar um arquivo Word sem a mensagem final de apoio do Time SSMAQ.

## Como abrir

1. Dê dois cliques em `abrir_app.bat`.
2. Escolha a classificação e preencha os campos do informativo.
3. Use `Copiar para WhatsApp` para copiar o texto completo.
4. Use `Gerar Word` para salvar o documento `.docx`.

## Organização

O arquivo `app.py` apenas inicia o sistema. A lógica principal fica na pasta `informativo_app`:

- `config.py`: campos e textos fixos.
- `text_builder.py`: montagem dos textos do WhatsApp e do Word.
- `docx_writer.py`: geração do arquivo `.docx`.
- `ui.py`: tela do aplicativo.

## Observação

O arquivo Word gerado não inclui a frase:

`👥 Time SSMAQ, solicitamos apoio para indicação da classificação do ocorrido com base nas informações acima (Ativo A1 / A2 / A3 ou Informativo).`

A seção `Documentos Anexos` aparece logo abaixo da descrição no corpo do documento.

O cabeçalho do Word usa a classificação marcada na tela, por exemplo:

`VAZAMENTO – 21/05/2026 – MAURILHO CEZAR BONFA - RVQ5G82 – CLAROS`
