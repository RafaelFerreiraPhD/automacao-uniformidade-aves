# Automação Zootécnica: Monitoramento de Uniformidade e Crescimento de Aves

Ferramenta desenvolvida em **Microsoft Excel e VBA** para automatizar a pesagem semanal, validação de regras de amostragem em galpão, cálculo biométrico de uniformidade e comparação com curvas de crescimento de linhagens comerciais.

---

## Por que esta planilha foi criada?

Durante a minha experiência técnica na produção comercial de postura, acompanhei semanalmente o desenvolvimento e o crescimento de pintainhas, realizando a pesagem e a avaliação da uniformidade dos lotes ao longo de toda a fase de cria e recria.

Na rotina de campo, identifiquei a ausência de uma ferramenta prática que integrasse três necessidades fundamentais:

* **Validação e cálculo imediato:** Reduzir o tempo gasto no processamento manual das pesagens e no cálculo dos índices de uniformidade ($\pm 10\%$).
* **Comparação direta com o padrão genético:** Cruzar o peso real das aves com as curvas oficiais dos manuais de linhagens comerciais de forma automática.
* **Rastreabilidade e série temporal:** Manter um histórico acumulado das pesagens anteriores, permitindo avaliar não apenas o lote ativo, mas a evolução temporal do galpão entre ciclos.

Esta aplicação foi desenvolvida exatamente para suprir esse gargalo operacional, transformando anotações isoladas num banco de dados estruturado para suporte técnico e tomada de decisão rápida.

---

## Demonstração da Planilha

### Menu Principal & Navegação
<p align="center">
  <img src="./img/inicio_ferramentas.png" alt="Tela Inicial e Ferramentas" width="850">
</p>

### Painel de Controle (Resultados e Curvas)
<p align="center">
  <img src="./img/painel_controle.png" alt="Painel de Controle" width="850">
</p>

### Coleta e Validação de Pesagens
<p align="center">
  <img src="./img/nova_coleta.png" alt="Nova Coleta de Pesagem" width="850">
</p>

---

## O que a ferramenta faz

1. **Cadastro do Lote:**
   - Registra data de alojamento, categoria (corte ou postura) e linhagem[cite: 8, 12]. O sistema vincula automaticamente as metas de peso para cada idade[cite: 6, 8].
2. **Validação de Amostragem na Coleta:**
   - Exige um mínimo de 100 aves ou 1% do galpão para validar a pesagem[cite: 3, 6, 8].
   - Impede o salvamento se faltar data ou ID do lote[cite: 2, 10].
3. **Cálculos Zootécnicos Automáticos:**
   - **Peso Médio e Desvio Padrão:** base do comportamento estatístico do galpão[cite: 2, 10].
   - **Uniformidade do Lote (%):** proporção de aves situadas na faixa de tolerância de $\pm 10\%$ em torno da média[cite: 2, 3, 8]:

$$\text{Uniformidade (\%)} = \left( \frac{\text{Aves com peso entre } 0{,}9 \times \bar{P} \text{ e } 1{,}1 \times \bar{P}}{\text{Total de aves pesadas}} \right) \times 100$$

4. **Classificação do Lote:**
   - Diagnóstico imediato por faixas de corte: Excelente ($\ge 85\%$), Boa ($\ge 80\%$), Regular ($\ge 70\%$) ou Crítica ($< 70\%$)[cite: 3, 7].
5. **Navegação em Tela Única:**
   - Transição entre telas controlada por rotinas VBA[cite: 3, 10]. As abas secundárias permanecem ocultas (`xlSheetVeryHidden`) para proteger bases de dados e fórmulas contra alterações acidentais[cite: 2, 5, 10, 12].
6. **Histórico Acumulado:**
   - As pesagens salvas alimentam uma base contínua com formatação padronizada, estruturada para auditorias técnicas ou conexão com ferramentas de Business Intelligence[cite: 2, 10].

---

## Organização dos Scripts VBA (Pasta `/src`)

Para consulta do código-fonte diretamente no repositório, as macros estão estruturadas em:

* **[`src/GestaoLotesEPesagens.bas`](./src/GestaoLotesEPesagens.bas):** cadastro de lotes e validação de pesagens semanais[cite: 2, 12].
* **[`src/NavegacaoSistema.bas`](./src/NavegacaoSistema.bas):** controle de navegação e visibilidade dinâmica das abas[cite: 5, 10].
* **[`src/ConfiguracoesEAdmin.bas`](./src/ConfiguracoesEAdmin.bas):** layout da tela de parâmetros e rotina de sanitização com senha de segurança[cite: 3, 11].

---

## Instruções de Uso

1. Baixe o arquivo binário na pasta [`app/Automacao_Uniformidade_Aves.xlsm`](./app/Automacao_Uniformidade_Aves.xlsm).
2. Abra a pasta de trabalho no Microsoft Excel e selecione **Habilitar Conteúdo** (Habilitar Macros).
3. Na tela de início, acesse **NOVO LOTE** para registrar as informações do lote alojado[cite: 8, 10].
4. Em **NOVA COLETA**, informe os pesos individuais das aves na coluna E[cite: 2, 8, 10].
5. Clique em **CHECAR RESULTADOS E SALVAR** para consolidar os dados no histórico e atualizar os gráficos no **PAINEL DE CONTROLE**[cite: 2, 8].
