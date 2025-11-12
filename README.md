# 🧩 Case Brasil Center — Governança de Produtos

> Projeto desenvolvido por **Bertil Soares** para demonstrar automação e análise de governança de produtos utilizando **Excel + Power Query + VBA**.

---

## 🎯 **Objetivo do Projeto**

Criar um **painel automatizado de governança de produtos**, com foco em **vigência, precificação e qualidade cadastral**.  
O objetivo é permitir que analistas visualizem, em tempo real:
- Quantos produtos estão **ativos ou expirados**;
- Quais expiram em até **7 dias**;
- A **distribuição por categoria, plano e tipo de oferta**;
- E quais produtos possuem **maior valor comercial**.

---

## ⚙️ **Arquitetura da Solução**

| Camada | Tecnologia | Descrição |
|--------|-------------|-----------|
| **Fonte de Dados** | CSV exportável de ERP | Base com cadastro, vigência e preços dos produtos |
| **Transformação** | Power Query | Tipagem, normalização e cálculo de status (`Ativo`/`Expirado`) |
| **Automação** | VBA (`Workbook_Open`) | Configuração dinâmica e atualização automática |
| **Visualização** | Excel Dashboard | Indicadores, gráficos e segmentações interativas |

---

## 🧠 **Principais Recursos**

### 🔹 Automacão VBA
O evento `Workbook_Open`:
- Detecta se o arquivo está sendo aberto pela primeira vez;
- Solicita o arquivo CSV base via seletor;
- Salva o caminho em aba oculta (`Configurações`);
- Atualiza as consultas Power Query automaticamente;
- Redireciona o usuário para o painel `Painel_Resumo`.

📘 *A aba “Configurações” é invisível via `xlSheetVeryHidden`, garantindo portabilidade e proteção.*

---

### 🔹 Fórmula de Status de Vigência

```excel
=SE(HOJE() > [@[Vigência_Fim]]; "Expirado"; "Ativo")
```

📌 *Usada para calcular o status de cada produto e alimentar os indicadores principais.*

Complementar:
```excel
=[@[Vigência_Fim]] - HOJE()
```
Define os **dias restantes** até o vencimento, permitindo alertas automáticos (produtos que expiram em 7 dias).

---

## 📊 **Indicadores e Métricas**

| Métrica | Descrição |
|----------|------------|
| **Total de produtos ativos** | Quantidade com status “Ativo” |
| **Total de produtos expirados** | Quantidade com status “Expirado” |
| **Expiram em 7 dias** | Produtos prestes a vencer |
| **Preço médio dos ativos** | Média dos produtos válidos |
| **Top 5 produtos por preço** | Priorizacão comercial |
| **Distribuição por categoria** | Internet, Telefonia, TV, Combo |

---

## 📈 **Design e Identidade Visual**
Cores inspiradas na paleta institucional da Brasil Center:

```
#AFBA40   #37BC7A   #125797   #EC0A0A
```

Visual com foco em **clareza operacional**, **alertas visuais automáticos** e **leitura executiva**.

---

## 🧾 **Documentação**
- [📘 Apresentação de Dados (PDF)](docs/2025.11.12%20-%20Case%20Brasil%20Center%20-%20Apresentação%20de%20Dados.pdf)
- [🧰 Documentação Técnica (PDF)](docs/2025.11.12%20-%20Case%20Brasil%20Center%20-%20Documentação%20Técnica.pdf)

Ambos detalham:
- Objetivo e arquitetura do case  
- Fluxo da automação VBA  
- Fórmulas aplicadas  
- Boas práticas técnicas  

---

## 🧭 **Fluxo de Automação**

```
[Abre Excel]
      ↓
[Verifica caminho CSV]
      ↓
┌────────────┬────────────┐
│ Caminho vazio│ Caminho salvo│
│ → Solicita   │ → Atualiza PQ│
│ → Salva CSV  │ → Mostra painel│
└────────────┴────────────┘
      ↓
[Exibe Painel_Resumo]
```

---

## 💡 **Destaques Técnicos**

✅ Conexão dinâmica (Power Query + VBA)  
✅ Portabilidade entre máquinas  
✅ Atualização automática  
✅ Painel visual e responsivo  
✅ Governança de vigência e precificação  

---

## 👨🏻‍💻 **Autor**

**Bertil Gonçalves Soares**  
📍 São Paulo — SP  
📧 [bertiljunior@gmail.com](mailto:bertiljunior@gmail.com)  
🔗 [linkedin.com/in/bertil-soares](https://linkedin.com/in/bertil-soares)  
💻 [github.com/BertilKenjiro](https://github.com/BertilKenjiro)

---

## 🧩 **Licença**
Uso livre para fins educacionais e portfólio.  
© 2025 Bertil Soares — Todos os direitos reservados.


