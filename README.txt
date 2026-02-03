DASHBOARD ESCOLA (MOBILE/PWA) – Conectado ao Google Sheets

1) Prepare o Google Sheets
- Abra sua planilha no Google Sheets.
- Crie uma aba chamada: DADOS_API
- Nessa aba, crie 2 colunas:
    A: chave
    B: valor

Chaves esperadas (recomendadas):
ticket_medio | 370
perda        | 0.30
custo_aluno  | 220
alunos_turma | 35
turmas_a     | (número do mês)
turmas_b     | (número do mês)
turmas_c     | (número do mês)
ocupacao     | 0.80   (80% = 0.80)
min_alunos   | 28     (menor turma do mês)
lucro_turma  | 1365
lucro_total  | 103740
cen_70       | 90090
cen_80       | 103740
cen_90       | 117390

2) Compartilhe / publique
Opção A (mais simples): Compartilhar -> "Qualquer pessoa com link - Leitor"
Opção B (mais robusta): Arquivo -> Compartilhar -> Publicar na Web -> (Planilha inteira)

3) Pegue o SHEET_ID
- Abra o Google Sheets.
- A URL fica assim:
  https://docs.google.com/spreadsheets/d/SHEET_ID/edit#gid=0
- Copie o trecho SHEET_ID (entre /d/ e /edit)

4) Configure o app
- Abra o arquivo app.js
- Troque:
  SHEET_ID: "COLE_AQUI_O_SHEET_ID"
  por seu ID real.

5) Usar no celular (PWA)
- Hospede estes arquivos em qualquer lugar (Drive, Github Pages, servidor, etc.).
- No iPhone/Android: Abrir no navegador -> "Adicionar à tela inicial".

Observação:
Se você abrir localmente (arquivo .html no celular), alguns navegadores bloqueiam fetch por segurança.
O ideal é hospedar (mesmo que seja em um servidor simples).

Power BI Mobile (resumo):
- Use a mesma aba DADOS_API como fonte (Excel/Sheets).
- Crie páginas: Geral / Operacional / Cenários.
- Configure alertas com cartões e semáforos usando medidas DAX.
