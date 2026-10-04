# 🐔 Sistema de Automação Zootécnica: Uniformidade de Lotes e Curvas de Crescimento

Aplicação em **Microsoft Excel e VBA** concebida para automatizar a gestão biométrica de aves (corte e postura), validação de regras de amostragem em galpão e acompanhamento de curvas padrão de linhagens comerciais.

---

## 🎯 Contexto e Desafio Operacional
A dispersão de pesos num lote avícola tem impacto direto na conversão alimentar, no rendimento de carcaça e na pontualidade do pico de postura. O processo tradicional de pesagem em papel ou em folhas de cálculo sem validação manual acarreta atrasos de cálculo e omissão de amostras representativas. 

Esta aplicação atua como um sistema integrado que valida as amostras recolhidas, calcula os índices biométricos instantaneamente e orienta a tomada de decisão técnica.

---

## 📸 Demonstração do Sistema

| Painel de Controlo & Desempenho | Recolha e Validação de Pesagens |
| :---: | :---: |
| ![Painel de Controlo](img/painel_controle.png) | ![Registo de Pesagem](img/nova_coleta.png) |

---

## ⚙️ Principais Funcionalidades

1. **Registo Controlado de Lote:**
   - Registo inicial do lote (identificação, linhagem, categoria e número de aves alojadas) com indexação automática às tabelas de referência técnica.
2. **Auditoria Biométrica na Recolha:**
   - Verificação em tempo real da regra de amostragem (mínimo de 100 aves ou 1% da população total).
   - Bloqueio de inserções incompletas ou datas inválidas.
3. **Métricas Zootécnicas Calculadas:**
   - **Peso Médio e Desvio Padrão ($\sigma$)**
   - **Coeficiente de Variação (CV%)**
   - **Percentagem de Uniformidade** na tolerância configurável (padrão $\pm 10\%$ da média):
     $$\text{Uniformidade (\%)} = \left( \frac{\text{Aves no intervalo } [0.9 \cdot \bar{P},\, 1.1 \cdot \bar{P}]}{\text{Total de aves pesadas}} \right) \times 100$$
4. **Arquitetura de Navegação Dinâmica:**
   - Utilização do estado `xlSheetVeryHidden` para apresentar apenas o ecrã ativo, protegendo as fórmulas e bases de dados contra edições indevidas.
5. **Base Histórica Estruturada:**
   - Gravação cronológica com tipagem e formatação numérica rigorosa para futuras auditorias e integração com ferramentas de Business Intelligence (Power BI).

---

## 📂 Organização dos Scripts VBA

O código fonte está modularizado na pasta [`/src`](./src):
* **[`src/GestaoLotesEPesagens.bas`](./src/GestaoLotesEPesagens.bas):** Rotinas `SalvarNovoLote()` e `ChecarResultadosESalvar()` — gestão de fluxos de entrada e validações de dados.
* **[`src/NavegacaoSistema.bas`](./src/NavegacaoSistema.bas):** Rotina `NavegarPara()` — controlo de interface e transições de ecrã.
* **[`src/ConfiguracoesEAdmin.bas`](./src/ConfiguracoesEAdmin.bas):** Construção dinâmica de layout e módulo `ResetarSistema()` com autenticação por palavra-passe.

---

## 🚀 Como Utilizar
1. Descarregue a folha de cálculo disponível em [`app/Automacao_Uniformidade_Aves.xlsm`](./app/).
2. Abra o ficheiro no Microsoft Excel e clique em **Habilitar Macros**.
3. No menu principal, inicie pelo registo do lote na aba **NOVO LOTE**.
4. Prossiga para a inserção das pesagens semanais na aba **NOVA COLETA**.
5. Clique em **CHECAR RESULTADO E SALVAR** para atualizar os gráficos do **PAINEL DE CONTROLE**.
