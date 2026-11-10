# Calculadora de Ganho Financeiro XP | M Wealth

Aplicativo em **Python + Streamlit** para apuração gerencial mensal do **ganho financeiro em reais** e da **rentabilidade estimada (Modified Dietz)** de contas XP, sem necessidade de precificar cada ativo, registrar PL diariamente ou aguardar o relatório de performance da corretora.

## 1. O que o aplicativo entrega

- Resultado individual por conta e consolidado das contas com dados completos.
- **Ganho financeiro** ajustado pelos aportes e retiradas externas.
- **Rentabilidade estimada** por Modified Dietz, ponderando os fluxos externos pelas datas.
- Revisão manual da classificação das movimentações diretamente na tela.
- Exportação para Excel (`.xlsx`) com resultado por conta, memória de movimentações, controle de PL, conferências e metodologia.

> **Importante:** é uma apuração gerencial. O ganho financeiro depende de PLs consistentes, extrato completo e classificação correta dos fluxos. Modified Dietz é uma aproximação da rentabilidade, **não** um TWR exato nem necessariamente a mesma rentabilidade informada pela XP.

## 2. Arquivos necessários

**Arquivos recorrentes: apenas dois.**

### A) Controle de PL Fechamento.xlsx

Mantenha uma única planilha histórica com estas colunas:

| Conta | Data | PL |
|---|---|---|
| 5166121 | 30/09/2026 | 13022474,24 |
| 5166121 | 30/10/2026 | *PL do fechamento de outubro* |
| 4163084 | 30/09/2026 | 8261890,04 |
| 4163084 | 30/10/2026 | *PL do fechamento de outubro* |

- `Conta`: número da conta XP, sempre no mesmo formato.
- `Data`: **data-base do PL**, não a data em que o arquivo foi emitido ou consultado.
- `PL`: patrimônio total referente àquele fechamento, com o mesmo critério de avaliação nas duas pontas (não misture PL bruto e líquido).
- Registre **somente o PL de fechamento mensal** de cada conta. Não é necessário atualizar diariamente.
- Evite registros duplicados para a mesma combinação de conta e data.
- Se o fim do mês cair em fim de semana, use o último dia útil de fechamento. Atenção: o aplicativo considera fins de semana, **mas não possui calendário de feriados**.

### B) Extrato XP do mês (.xlsx)

O aplicativo espera colunas com os seguintes nomes:

- `CLIENTE` (número da conta)
- `DATA DE REFERENCIA`
- `DATA DE LIQUIDACAO`
- `DESCRICAO`
- `VALOR`
- `SALDO` (opcional)

É possível importar um ou mais extratos. O extrato precisa cobrir o mês de referência integralmente e trazer os fluxos externos relevantes. **Não importe arquivos com os mesmos lançamentos repetidos:** o aplicativo concatena os extratos, mas não deduplica operações.

**Implantação (opcional):** se o controle ainda não tiver os PLs de abertura, é possível enviar as planilhas de `Posição Consolidada` das contas e indicar sua data-base no campo de implantação. Depois disso, a rotina normal requer apenas os dois arquivos acima.

## 3. Instalação e execução

Recomendado: **Python 3.10 ou superior**.

Coloque na mesma pasta:

```text
conciliacao-xp/
├── app_conciliacao_xp.py
├── requirements_conciliacao_xp.txt
└── README_conciliacao_xp.md
```

Abra o terminal nessa pasta e execute:

```bash
python -m pip install -r requirements_conciliacao_xp.txt
python -m streamlit run app_conciliacao_xp.py
```

No Windows, se o comando `python` não funcionar, tente `py` no lugar de `python`.

O Streamlit normalmente abre o aplicativo no navegador. Caso contrário, use o endereço local informado pelo terminal (em geral `http://localhost:8501`).

## 4. Fluxo de uso mensal

1. Acrescente ao **Controle de PL** o fechamento das cinco contas no fim do mês.
2. Exporte o **extrato completo** do mesmo mês na XP.
3. Abra o aplicativo e selecione o **mês de referência** (a escolha do dia no calendário não muda o mês calculado).
4. Selecione a data a ser usada para classificar os fluxos: `Data de referência` ou `Data de liquidação`. Use aquela que representa a entrada ou saída efetiva do PL; confira movimentações na virada do mês.
5. Carregue o Controle de PL e o(s) extrato(s). Se for a primeira execução e faltarem PLs de abertura, utilize a opção de implantação.
6. Abra **Conferência das movimentações** e revise principalmente TEDs, PIX, transferências e itens classificados como `Revisar`.
7. Confira PL inicial/final, fluxo externo líquido, ganho financeiro, rentabilidade e alertas.
8. Baixe o Excel em **Baixar apuração consolidada e memória de cálculo** e arquive-o com os arquivos de origem.

Para apurar **novembro/2026**, por exemplo, o aplicativo utiliza os PLs de fechamento de outubro como abertura e os PLs de fechamento de novembro como encerramento, acompanhados do extrato de novembro.

## 5. Metodologia

### Ganho financeiro (em R$)

```text
Fluxo externo líquido = aportes externos - retiradas externas
Ganho financeiro = PL final - PL inicial - fluxo externo líquido
```

Convenção do extrato: entradas externas com sinal positivo; saídas externas com sinal negativo. Aplicações e resgates **dentro da mesma conta** não são fluxos externos para este cálculo.

### Rentabilidade estimada (Modified Dietz)

```text
Capital ponderado = PL inicial + soma(fluxo externo × peso temporal)
Peso temporal = dias restantes até o fim do mês / dias do período
Rentabilidade estimada = ganho financeiro / capital ponderado
```

O resultado percentual é exibido somente quando o capital ponderado é positivo. Os pesos consideram as datas das movimentações selecionadas e o último dia **calendário** do mês; se o último pregão for anterior, pode haver pequena diferença metodológica.

### Transferências entre contas do mesmo grupo

- Para a **conta individual**, uma transferência que efetivamente sai do PL deve ser tratada como **retirada externa daquela conta**, e a entrada correspondente como **aporte externo na outra conta**.
- Para o **consolidado**, essas transferências devem se compensar quando ambas as pontas estiverem no período e no universo de contas. Confira se não há lançamentos ausentes ou diferenças de data.
- Aplicações e resgates de produtos *dentro da mesma conta* permanecem como movimentações internas.

## 6. Validações antes de entregar ao condomínio

- [ ] PL inicial e final de **todas** as contas do grupo disponíveis na data-base correta.
- [ ] Extrato completo e sem duplicações para o período escolhido.
- [ ] TEDs, PIX e transferências revisados; nenhum lançamento permanece como `Revisar`.
- [ ] Nenhuma movimentação tem sinal incompatível com sua classificação.
- [ ] PLs inicial e final usam critérios patrimoniais comparáveis.
- [ ] Transferências entre contas do grupo foram conciliadas nas duas pontas.
- [ ] Não há movimentações relevantes pendentes de liquidação ou faltantes no extrato.

**Atenção:** o aplicativo pode exibir e exportar números mesmo com avisos de pendência. Não trate esses resultados como definitivos sem resolver as inconsistências. A ausência de lançamentos no fim do mês não comprova, por si só, que o extrato esteja completo.

## 7. Relatório Excel de saída

O arquivo exportado se chama `conciliacao_xp_AAAA_MM.xlsx` e contém:

| Aba | Conteúdo |
|---|---|
| Resultado por conta | PL inicial e final, fluxo líquido, ganho financeiro, Modified Dietz e status |
| Memória movimentações | Linhas do extrato, classificação, fluxos externos, pesos e cálculos |
| Controle PL | Histórico de PL utilizado |
| Conferências | Pendências e datas de cobertura dos extratos |
| Metodologia | Critérios e alertas da apuração |

O consolidado soma **somente contas com ambos os PLs disponíveis**. Se alguma conta estiver incompleta, não interprete essa soma como resultado integral do grupo.

## 8. Limitações conhecidas

- Não precifica ativos individualmente e não reconstrói PL final a partir de operações: o **PL final deve ser informado**.
- Não calcula TWR exato: a métrica percentual utilizada é **Modified Dietz**.
- Não deduplica extratos sobrepostos.
- Não valida automaticamente feriados nacionais, estaduais ou municipais.
- Não garante que o extrato esteja completo: essa conferência ainda é operacional.
- As edições de classificação são feitas para a **apuração corrente** e aparecem na exportação; não constituem cadastro persistente de regras por cliente.
- O resultado é gerencial; não substitui documentação contábil, fiscal ou relatório oficial da XP.

## 9. Solução de problemas

**`ModuleNotFoundError`**: instale as dependências com `python -m pip install -r requirements_conciliacao_xp.txt` usando o mesmo ambiente Python que executa o Streamlit.

**`Controle de PL deve conter as colunas Conta, Data e PL`**: confira o cabeçalho da primeira aba da planilha.

**`Extrato ... sem colunas`**: verifique se foi exportado no mesmo layout XP esperado pelo aplicativo.

**`PL inicial/final ausente`**: confira a data-base e as contas do Controle de PL. O aplicativo busca o último dia útil do fechamento anterior e do mês de referência (considerando fins de semana).

**Movimentos classificados como `Revisar`**: abra o editor de movimentações e escolha `Aporte externo`, `Retirada externa` ou `Interna / investimento`, conforme o efeito efetivo no PL.

**Ganho financeiro inesperado**: confira sinais de TEDs, transferências entre contas, extrato completo, eventuais linhas duplicadas e critérios de PL.

---

**Uso interno — M Wealth.** Confira os números antes de compartilhar a apuração com o cliente ou com o financeiro do condomínio.
