# Regra global de escrita da memória compartilhada

Status: documentação operacional; implantação automática em cada agente requer teste independente.
Fonte canônica: Google Drive `Cloud/MD/Memoria-Compartilhada/`.
Protocolo oficial: `protocolos/SINCRONIZACAO_V2.md` (versão 3, revisão de 2026-11-13).
Índice: https://drive.google.com/file/d/1k3QK5bbN6HkB8LrQDeZe1WoQ7Gu7LTzn/view

## Procedimento obrigatório para ChatGPT, Codex, Claude e Claude Code

1. Antes de escrever, consultar INDEX, protocolo V2, mestre STE-AI e contexto pertinente; para o Projeto Claudia consultar também V3 e considerar as planilhas oficiais a fonte da verdade.
2. Identificar arquivo canônico por ID e pasta; ler conteúdo, versão e modifiedTime; detectar alterações concorrentes e parar em caso de conflito.
3. Preservar backup da versão anterior em `historico/` e confirmar a gravação.
4. Para arquivos Markdown existentes no Drive, atualizar conteúdo no **mesmo ID** por Drive `files.update` ou conector `update_file` com `file_uri` e MIME `text/markdown`. Não usar upload com `overwrite=true` via Library para Drive: foi observado erro `overwrite_not_supported`.
5. Reler o mesmo ID e validar conteúdo, versão e integridade. Não presumir que uma checagem anterior impede corrida entre leitura e gravação.
6. Para fatos duráveis, criar um evento **novo e imutável** em `eventos/`, com origem, data/hora verificável, recurso, resumo, estado e verificação; reler para confirmar.
7. Editar Google Docs, Sheets e Slides nativos pelas APIs próprias, nunca convertendo-os implicitamente em Markdown.
8. Sem operação segura ou acesso, manter o original, registrar pendência e informar `não executado` ou `falhou`; não criar duplicata concorrente nem alegar sincronização.
9. Atualizar fluxo, índice, versão e histórico quando o projeto for alterado ou republicado; conferir com o estado real antes de declarar ativo.

Esta página é um apontador para o protocolo oficial, não uma segunda fonte de verdade. Sua presença no repositório não prova que o agente a executa automaticamente.
