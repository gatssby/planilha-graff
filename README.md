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

No topo há um resumo compacto com:

- Pessoa
- Situação
- Valor
- Balanço visual

Logo abaixo ficam os quatro campos de entrada:

1. Pessoa
2. Data
3. Valor
4. Observação

## Nova regra de lançamento

Todo valor lançado representa um valor que aquela pessoa pagou em favor dos avós.

Por isso:

- selecione sempre quem pagou;
- informe o valor sempre como número positivo;
- use Observação para indicar o motivo, por exemplo: Remédio, Consulta, Mercado, Transporte etc.

Exemplo:

- `Valciria | 06/10/2026 | 250 | Remédios`
- `Deco | 07/10/2026 | 180 | Consulta`

Não existem mais lançamentos negativos.

## Lógica estilo Splitwise

A planilha soma tudo que foi pago pelos cinco irmãos e calcula a média:

`parcela de referência = total pago ÷ 5`

Depois compara quanto cada pessoa pagou com essa parcela:

`saldo individual = total pago pela pessoa − parcela de referência`

A situação aparece assim:

- **Deve contribuir R$ X**: a pessoa pagou menos do que a média dos cinco naquele momento.
- **A haver R$ X**: a pessoa pagou mais do que a média.
- **Em dia**: a pessoa está exatamente na média.

Ninguém deve diretamente para outra pessoa. O número funciona apenas como balanço de contribuição entre os cinco.

## Balanço visual

A coluna **Balanço** usa uma pequena barra com zero no centro:

- vermelho para a esquerda = abaixo da média;
- verde para a direita = acima da média;
- cinza = em dia.

A intensidade da barra é relativa ao maior desequilíbrio atual entre os cinco. O valor exato continua aparecendo na coluna **Valor**.

## Aplicar ou atualizar

1. Abra a planilha no Google Sheets.
2. Vá em **Extensões → Apps Script**.
3. Abra `Code.gs`.
4. Substitua todo o conteúdo pelo `Code.gs` deste repositório.
5. Salve.
6. No seletor de funções do Apps Script, escolha `setupSpreadsheet`.
7. Clique em **Executar** e autorize, se solicitado.

A versão atual tenta preservar os lançamentos já existentes antes de reconstruir o layout.

**Importante:** lançamentos antigos negativos pertencem à lógica anterior e não entram mais nos cálculos. Se houver dados reais antigos, ajuste-os para o novo formato: escolha quem pagou e use o valor positivo.

Depois disso, qualquer alteração nos lançamentos atualiza o resumo automaticamente.
