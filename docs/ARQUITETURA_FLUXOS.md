# Arquitetura e fluxos — Central de Eventos

Documentação 1.0 · 13/11/2026 · [HTML interativo](fluxos.html)

## Etapa 1 · Inventário

Base GitHub `99e89e026423ca86819e05d382383a64b8af125e`.

| Componente | Entrada e ação | Fonte |
| --- | --- | --- |
| Planilha | Linhas de Eventos e calendário escolhido. O operador usa os comandos da planilha. A agenda é resolvida por linha. | [Utils.gs](https://github.com/duduxweb/central-de-eventos-planilhas/blob/99e89e026423ca86819e05d382383a64b8af125e/Utils.gs) |
| Apps Script | Valida a seleção e resolve a agenda. O onEdit só registra que a sincronização automática está desativada. O código contém variantes de sincronização por seleção. | [selecionado.gs](https://github.com/duduxweb/central-de-eventos-planilhas/blob/99e89e026423ca86819e05d382383a64b8af125e/selecionado.gs) |
| Calendar | Cria ou atualiza eventos pelo comando autorizado. A permissão da conta executora afeta a operação. IDs e status permitem acompanhar a associação com as linhas. | [AgendasTools.gs](https://github.com/duduxweb/central-de-eventos-planilhas/blob/99e89e026423ca86819e05d382383a64b8af125e/AgendasTools.gs) |

## Etapa 2 · Evidências e pendências

Funções Apps Script e gatilho onEdit conferidos no GitHub. Checkout local identificado anteriormente, sem alterar arquivos ou calendários.

Work deve conferir o contêiner Apps Script, a versão instalada, os gatilhos e as permissões de calendário. Sem criar, mover ou excluir eventos de teste.

## Histórico

- 1.0 · Primeiro mapa interativo. Atualize o índice e o HTML na mesma revisão e confira a cópia do Drive.
