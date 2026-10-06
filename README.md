# Planilha Graff

Google Apps Script para criar uma planilha simples de controle familiar no Google Sheets, pensada principalmente para uso pelo celular.

## Pessoas

- Valciria
- Deco
- Vivi
- Nega
- Valdi

## Estrutura

Tudo fica na mesma aba **Lançamentos**.

No topo há somente um resumo compacto com a situação de cada pessoa. Logo abaixo ficam os quatro campos de entrada:

1. Pessoa
2. Data
3. Valor
4. Observação

Não há meta, poupança ou barra de progresso.

## Lógica estilo Splitwise

A planilha não diz que uma pessoa deve dinheiro para outra.

Ela calcula quanto cada pessoa já contribuiu em comparação com sua parcela dos gastos realizados:

`saldo individual = contribuições da pessoa − (gastos totais ÷ 5)`

A situação aparece assim:

- **Deve contribuir R$ X**: a pessoa ainda está abaixo da sua parcela atual dos gastos e precisa contribuir esse valor para ficar em dia.
- **A haver R$ X**: a pessoa já colocou mais dinheiro do que sua parcela dos gastos. Esse crédito continua atribuído a ela.
- **Em dia**: a pessoa está exatamente equilibrada naquele momento.

Isso evita a lógica de “X deve para Y” e mantém o acerto em torno do caixa comum da família.

## Regra de lançamento

- **Contribuição:** selecione a pessoa e informe um valor positivo.
- **Gasto com os avós:** informe um valor negativo. O campo Pessoa pode ficar em branco.

Exemplos:

- `Valciria | 06/10/2026 | 500 | Contribuição`
- ` | 06/10/2026 | -200 | Remédio`

## Aplicar ou atualizar

1. Abra a planilha no Google Sheets.
2. Vá em **Extensões → Apps Script**.
3. Abra `Code.gs`.
4. Substitua todo o conteúdo pelo `Code.gs` deste repositório.
5. Salve.
6. No seletor de funções do Apps Script, escolha `setupSpreadsheet`.
7. Clique em **Executar** e autorize, se solicitado.

A versão atual tenta preservar os lançamentos já existentes antes de reconstruir o layout. Mesmo assim, para uma planilha que já esteja em uso real, é recomendável fazer uma cópia de segurança antes de executar `setupSpreadsheet()`.

Depois disso, as alterações feitas nos lançamentos atualizam o resumo automaticamente.
