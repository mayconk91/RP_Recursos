# Planejador de Recursos

Aplicativo web offline para cadastro de recursos e atividades, acompanhamento de timeline, indicadores de capacidade, exportacoes e backup.

## Como abrir

Abra `index.html` em um navegador moderno. Para usar recursos de PWA/offline com service worker, o ideal e servir a pasta por um servidor local simples em vez de abrir direto pelo arquivo.

## Arquivos principais

- `index.html`: tela principal e estrutura do aplicativo.
- `styles.css`: estilos principais da interface.
- `app.js`: regras de negocio, persistencia local, importacao/exportacao e renderizacao.
- `enhancer2.js`: melhorias carregadas pela pagina principal.
- `planner_enhancer.js`: complemento carregado pela pagina principal.
- `tour.js` e `tour.css`: guia rapido do aplicativo.
- `service-worker.js`: cache offline do app.
- `manifest.webmanifest`: configuracao PWA.
- `version.json`: versao atual publicada.
- `icons/`: icones do aplicativo.
- `docs/changelogs/`: historico de alteracoes antigas.

## Dados e backup

O app usa armazenamento local do navegador e tambem possui rotinas de backup/exportacao. Antes de alterar regras importantes, exporte um backup pela propria interface.

## Manutencao recomendada

- Criar controle de versao Git para acompanhar futuras alteracoes.
- Dividir `app.js` em modulos menores quando houver tempo para uma refatoracao com testes.
- Revisar arquivos legados `enhancer.js`, `enhancer.css` e `capacidade_v4.html`, que aparecem no cache offline mas nao sao carregados diretamente por `index.html`.
- Manter `VERSION` em `service-worker.js` sincronizada com `version.json` sempre que publicar uma nova versao.
