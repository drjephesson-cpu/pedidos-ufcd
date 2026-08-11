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

### Histórico e Neon

Na Vercel o disco é temporário. Para **guardar pedidos por data**, conecte o Neon:

1. No [Neon Console](https://console.neon.tech), crie um projeto e copie a connection string.
2. Na Vercel → Project **ufcd** → **Settings → Environment Variables**, adicione:
   - `DATABASE_URL` = `postgresql://...?...sslmode=require`
   - `SECRET_KEY` = uma chave aleatória
3. Faça Redeploy.

Localmente, copie `.env.example` para `.env` e preencha o mesmo `DATABASE_URL`. Sem Neon, o histórico usa SQLite em `data/pedidos.db` (só na máquina local).
