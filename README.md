# 📊 CryptoDominance Tracker

---

## 🧠 Sobre o projeto

Durante minhas férias, com um tempo livre a mais (e um pouco de tédio 😅), resolvi estudar APIs de criptomoedas e criar algo prático.  
Este projeto busca automaticamente dados atualizados das principais criptomoedas do mundo, gera uma planilha em Excel e permite que o Power BI atualize os gráficos sempre com os valores mais recentes.

---

## 🚀 O que o projeto faz

- Busca o **Top 10 criptomoedas por market cap**
- Coleta:
  - Preço atual (USD)
  - Market Cap
  - Variação em 24h
- Calcula a **dominância (%) de cada cripto no mercado global**
- Gera automaticamente um arquivo **Excel (.xlsx)**
- O **Power BI consome esse arquivo** e atualiza os dados com apenas um refresh

---

## 📁 Estrutura da planilha gerada

A planilha é salva na **Área de Trabalho do usuário** com o nome:

top10_cripto_usd.xlsx

Colunas geradas:

- DataColeta
- Rank
- Cripto
- Simbolo
- PrecoUSD
- MarketCapUSD
- DominanciaPct
- Variacao24h

Essa estrutura é ideal para:
- Gráficos
- Cards
- Rankings
- Treemaps
- Séries temporais (se quiser evoluir o projeto)

---

## 🔄 Integração com Power BI

O Power BI lê diretamente o arquivo Excel gerado pelo script.

Fluxo simples:
1. Executa o script Python
2. O Excel é atualizado/substituído
3. No Power BI, basta clicar em **Atualizar**
4. Todos os visuais refletem os novos valores automaticamente

Sem retrabalho, sem edição manual.

---

## 🛠️ Tecnologias usadas

- **Python 3**
- **API CoinGecko** (dados de mercado cripto)
- **requests** (requisições HTTP)
- **openpyxl** (geração de Excel)
- **Power BI** (visualização e análise)

---

## 📦 Dependências

Instale antes de rodar o projeto:

```bash
pip install requests openpyxl
```

## ▶️ Como executar
python crypto_top10_xlsx.py


Após a execução:

O arquivo Excel será atualizado automaticamente

O Power BI poderá ser atualizado com um clique

## 💡 Observações finais

Este projeto foi criado com foco em aprendizado, curiosidade e automação simples.
Ele pode ser facilmente expandido para histórico diário, alertas ou dashboards mais avançados.

Sinta-se à vontade para adaptar e evoluir 🚀


---
