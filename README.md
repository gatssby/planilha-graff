# Planilha Graff

Google Apps Script para criar uma planilha simples de controle familiar no Google Sheets, pensada principalmente para uso pelo celular.

## Pessoas

- Valciria
- Deco
- Vivi
- Nega
- Valdi

## Estrutura

Tudo fica na mesma aba **Lançamentos**:

- resumo financeiro e saldos individuais no topo;
- lançamentos logo abaixo;
- quatro campos de entrada: Pessoa, Data, Valor e Observação.

### Regra de lançamento

- **Contribuição:** selecione a pessoa e informe um valor positivo.
- **Gasto com os avós:** informe um valor negativo. Pessoa pode ficar em branco.

O resumo calcula automaticamente total contribuído, total gasto, reserva atual, progresso até R$ 3.000, cota dos gastos e saldo individual.

Saldo individual = contribuição da pessoa − 1/5 dos gastos realizados.

## Instalação / atualização

1. Abra a planilha no Google Sheets.
2. Vá em **Extensões → Apps Script**.
3. Substitua todo o conteúdo de `Code.gs` pelo `Code.gs` deste repositório.
4. Salve.
5. No seletor de funções do Apps Script, escolha `setupSpreadsheet`.
6. Clique em **Executar** e autorize, se solicitado.

**Atenção:** `setupSpreadsheet()` recria o layout da aba Lançamentos e limpa seu conteúdo. Faça isso antes de começar a lançar dados. Depois da configuração inicial, não é necessário executá-la novamente para uso normal.

A função também remove a antiga aba **Resumo**, caso exista, pois o resumo agora fica na própria aba **Lançamentos**.
