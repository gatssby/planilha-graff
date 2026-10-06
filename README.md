# Planilha Graff

Google Apps Script para criar uma planilha simples de controle familiar no Google Sheets, pensada principalmente para uso pelo celular.

## Pessoas

- Val Síria
- Deco
- Vivi
- Nega
- Valdi

## Como funciona

A aba **Lançamentos** tem somente quatro campos visíveis:

1. Pessoa
2. Data
3. Valor
4. Observação

### Regra de lançamento

- **Contribuição:** selecione a pessoa e informe um valor positivo.
- **Gasto com os avós:** informe um valor negativo. O campo Pessoa pode ficar em branco.

Exemplo:

- Val Síria contribuiu R$ 500 → `Val Síria | 06/10/2026 | 500 | Contribuição`
- Houve R$ 200 de gasto com remédio → `| 06/10/2026 | -200 | Remédio`

## Cálculos automáticos

O resumo calcula:

- total contribuído;
- total gasto;
- reserva atual = contribuições − gastos;
- meta inicial de reserva de R$ 3.000;
- barra de progresso da reserva;
- contribuição acumulada de cada pessoa;
- cota atual dos gastos = gastos totais ÷ 5;
- saldo individual = contribuição da pessoa − cota atual dos gastos.

Saldo positivo significa valor a haver daquela pessoa. Saldo negativo significa que, considerando apenas os gastos efetivamente registrados até aquele momento, ela contribuiu menos que sua parcela de 1/5.

O excedente de uma pessoa permanece atribuído a ela no saldo individual e, enquanto não for gasto, também compõe a reserva familiar.

## Instalação

1. Crie uma planilha vazia no Google Sheets.
2. Abra **Extensões → Apps Script**.
3. Substitua o conteúdo de `Code.gs` pelo arquivo deste repositório.
4. Se quiser, copie também o conteúdo de `appsscript.json` para o manifesto do projeto.
5. Salve e execute `setupSpreadsheet()` uma vez.
6. Autorize o script quando o Google solicitar.

Depois disso, use apenas a aba **Lançamentos** no dia a dia. A aba **Resumo** é atualizada automaticamente a cada edição.
