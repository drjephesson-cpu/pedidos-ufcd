# Pedidos UFCD — Farmácia

Site local de pedido de medicamentos com **catálogo fixo** (abas da planilha) e atualização diária do estoque AGHU.

## Como usar

```powershell
cd pedido-estoque
py -m pip install -r requirements.txt
py app.py
```

Abra: http://127.0.0.1:5000

Ou dê dois cliques em `iniciar.bat`.

### Fluxo

1. O catálogo já vem embutido em `data/catalogo.json` (363 itens / 11 abas).
2. Envie o **EstoqueFarmacia-*.xlsx** do dia — o match é pelo Cód. AGHU.
3. Navegue pelas abas, filtre “Só Pedir = Sim” e exporte o pedido.

### Lógica

- **Pedir?** quando estoque &lt; ponto de pedido  
- **Quanto Pedir?** = arredonda para cima até o estoque mínimo (múltiplo da caixa)  
- Coluna **Manual** sobrescreve a quantidade  

### Atualizar o catálogo (raro)

Se a planilha mestre mudar:

```powershell
py extrair_catalogo_fix.py
```

(Digite a senha do Excel se pedir.)
