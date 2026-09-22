# Painel de Indicadores de Sincronização

Sistema em Google Apps Script que acompanha a produtividade da equipe na sincronização diária de anúncios entre o **Bling** (ERP) e o e-commerce, com um painel web de indicadores.

## O problema

A equipe sincroniza anúncios todos os dias, mas não havia como acompanhar o ritmo nem medir o avanço em relação às metas. O controle era informal, e a gestão não tinha visibilidade do que era feito ao longo do dia.

## O que o sistema faz

1. **Registra automaticamente o horário** de cada sincronização. O colaborador só cola o SKU numa aba, e um gatilho `onEdit` grava o timestamp ao lado, sem ação manual.
2. **Identifica o produto-pai de cada SKU** consultando a API do Bling. Se o código for de uma variação, o sistema sobe até o produto-pai e remove duplicatas, para que cada produto conte uma vez só.
3. **Gera o relatório diário** com os produtos sincronizados e uma linha de resumo, e limpa a aba de registro para o dia seguinte.
4. **Mantém os indicadores atualizados:** total do dia, da semana e do ciclo mensal, distribuição por hora do expediente e percentual da meta mensal. Metas de 25/dia, 125/semana e 500/mês.
5. **Serve um painel web** (`doGet`) com cartões de KPI, medidor de meta em arco e três gráficos (área por hora, barras por dia da semana e por semana do ciclo), todos desenhados em SVG diretamente no navegador, sem bibliotecas externas.
6. **Reinicia os ciclos sozinho:** gatilhos semanais movem a semana atual para "anterior", zeram os contadores e reiniciam o ciclo mensal a cada quatro semanas.

## Desafios técnicos

- **Limite de requisições da API (HTTP 429):** implementei novas tentativas com espera crescente (*backoff* exponencial: 1s, 2s, 4s...), mais pausas entre chamadas. Sem isso, sincronizações eram perdidas silenciosamente.
- **Execuções simultâneas:** o `LockService` impede que duas execuções do relatório rodem ao mesmo tempo e gravem uma por cima da outra.
- **Hierarquia de produtos no Bling:** um SKU pode ser o produto-pai, uma variação ou um produto simples. A função de busca trata os três casos, seguindo a referência `variacao.produtoPai` quando existe.
- **Horários fora do expediente:** sincronizações antes das 8h ou depois das 17h eram descartadas do gráfico. Agora são agrupadas no horário mais próximo, em vez de sumirem.
- **Reset de ciclos sem intervenção:** o estado do ciclo (semana 1 a 4) fica na própria planilha e avança por gatilho de tempo, com o acumulado mensal zerando na virada.

## Tecnologias

Google Apps Script (JavaScript), API REST do Bling v3, OAuth2, Google Sheets, HTML Service, SVG, gatilhos automáticos.

## Como configurar

1. Em *Configurações do projeto > Propriedades do script*, cadastre `BLING_CLIENT_ID` e `BLING_CLIENT_SECRET`.
2. Preencha a constante `SPREADSHEET_ID` com o ID da sua planilha.
3. Rode `iniciarAutorizacao` e autorize o acesso ao Bling pelo link gerado.
4. Rode `inicializarAbaGraficos` uma vez para criar a estrutura de dados e os gráficos.
5. Rode `criarGatilhosGraficos` para agendar os resets semanais, e crie um gatilho diário para `gerarRelatorioDiario`.
6. Implante como aplicativo da web para acessar o painel.

## Próximas melhorias

- Substituir `openById` por `getActiveSpreadsheet()`, já que o projeto está vinculado à planilha. Isso elimina o ID fixo no código.
- Exportar o histórico para uma base própria, em vez de manter o estado apenas nas células da planilha.

---

*Regras de negócio, arquitetura e validação definidas por mim; código desenvolvido com assistência de IA.*
