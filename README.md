# Despacho App Utilities

Script para processar registros de "Abertura da tela de Despacho" e exportar para Excel.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

1. Prepare a data inicial do turno (formato `YYYY-MM-DD`).
2. Cole várias linhas no formato:
   ```
   <hh:mm> - Abertura da tela de Despacho - <EMP> - EXCEDIDO EM: <xx>%
   ```
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
