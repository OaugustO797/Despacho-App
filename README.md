# Despacho App Utilities

Script para processar registros de "Abertura da tela de Despacho" e exportar para Excel, acompanhado de um painel web simples para visualização e inserção manual.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso do utilitário em linha de comando

1. Prepare a data inicial do turno (formato `YYYY-MM-DD`).
2. Cole várias linhas no formato:
   ```
   <hh:mm> - Abertura da tela de Despacho - <EMP> - EXCEDIDO EM: <xx>%
   ```
   - Horários são aceitos como `00:00`, `00h00` ou `00:00h`.
3. Execute o utilitário apontando a data e opcionalmente um arquivo de saída:

```bash
python src/processor.py --data 2024-05-10 --output despacho.xlsx <<'EOF'
05:15 - Abertura da tela de Despacho - ABC - EXCEDIDO EM: 12%
23:50 - Abertura da tela de Despacho - XYZ - EXCEDIDO EM: 34%
00:10 - Abertura da tela de Despacho - XYZ - EXCEDIDO EM: 45%
EOF
```

- Mudanças de dia são detectadas quando o horário volta no registro seguinte.
- Horários exportados recebem +3h e sufixo `Z` para compatibilidade com PowerAutomate/SharePoint.
- Cada tabela da planilha comporta até 256 linhas; tabelas adicionais são criadas lado a lado na mesma aba.
- Em vez de colar no STDIN você pode usar `--input caminho/para/arquivo.txt`.

## Painel web

Abra `web/index.html` em um navegador ou sirva o diretório com `python -m http.server` e acesse `http://localhost:8000/web/`.

- Cadastre registros individualmente ou cole várias linhas seguindo o padrão original.
- A data do turno é obrigatória; a lógica detecta a troca de dia quando o horário volta após 23:59.
- Os botões CSV/TXT exportam com horário +3h e sufixo `Z`, adequados para PowerAutomate/SharePoint.
