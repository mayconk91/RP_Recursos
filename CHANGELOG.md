# Changelog

Historico consolidado a partir dos arquivos antigos em `docs/changelogs/`.

## v1.2.8.1

Planejador de Recursos v1.2.8.1

Ajustes controlados sobre a base revisada v1.2.8.0:
- alinhamento da versÃ£o interna do sistema, version.json e service worker para 1.2.8.1;
- bloqueio do salvamento quando a janela de concorrÃªncia nÃ£o estabiliza apÃ³s as tentativas previstas;
- bloqueio do salvamento quando o BD foi alterado por outra sessÃ£o desde a Ãºltima sincronizaÃ§Ã£o local;
- remoÃ§Ã£o da indicaÃ§Ã£o enganosa de usuÃ¡rio no overlay de espera;
- endurecimento do sistema de toast para montagem segura com textContent/createElement;
- remoÃ§Ã£o de datalist duplicado no index.html.

## v1.2.8.2

v1.2.8.2 - Ajustes de concorrÃªncia/UX controlados

- restaura o fluxo de fila de gravaÃ§Ã£o sem bloquear o usuÃ¡rio por alteraÃ§Ã£o externa genÃ©rica;
- remove a exigÃªncia de "sincronizar" quando nÃ£o existe esse botÃ£o na interface;
- preserva a ediÃ§Ã£o local durante a espera da fila de gravaÃ§Ã£o;
- atualiza o status do topo durante gravaÃ§Ã£o, erro e apÃ³s recarga automÃ¡tica do BD.

## v1.2.8.5

Planejador de Recursos v1.2.8.5

- restaura o bloqueio do save quando a fila de gravaÃ§Ã£o nÃ£o estabiliza apÃ³s o limite de tentativas;
- mantÃ©m o fluxo de prosseguir apÃ³s a espera normal da fila, sem tratar isso como conflito fatal;
- preserva o layout com status ao lado das abas e a correÃ§Ã£o do XSS em runAvailability;
- alinha a versÃ£o interna do app, version.json e service-worker para 1.2.8.5.

## v1.2.8.6

Planejador de Recursos v1.2.8.6

CorreÃ§Ã£o de sincronizaÃ§Ã£o entre sessÃµes para exclusÃµes e atualizaÃ§Ãµes externas:
- removida a reaplicaÃ§Ã£o indevida do eventLog local apÃ³s recarga externa do BD/FSA;
- o estado vindo do banco agora passa a ser adotado como baseline persistido da sessÃ£o;
- apÃ³s salvar no BD, o baseline local tambÃ©m Ã© renovado para evitar reprocessar eventos jÃ¡ persistidos;
- removido o salvamento automÃ¡tico indevido apÃ³s watcher do FSA detectar alteraÃ§Ã£o externa.

Impacto esperado:
- exclusÃµes feitas em outro dispositivo passam a sumir corretamente na sessÃ£o aberta, sem exigir sair e entrar novamente;
- reduz o risco de ressuscitar atividades excluÃ­das por reaplicaÃ§Ã£o de eventos antigos locais.


Complemento local: comentÃ¡rios em atividade + exclusÃ£o auditÃ¡vel com justificativa obrigatÃ³ria e trilha preservando comentÃ¡rio da atividade.


Complemento local:
- comentÃ¡rios de atividade passaram a aceitar mÃºltiplos registros
- cada novo comentÃ¡rio Ã© adicionado como registro imutÃ¡vel
- comentÃ¡rios publicados nÃ£o podem ser editados; apenas novos comentÃ¡rios podem ser incluÃ­dos
- persistÃªncia no banco apontado com coluna adicional comentariosJson para manter o histÃ³rico estruturado

## v1.2.8.7

Planejador de Recursos v1.2.8.7

Ajustes de continuidade para controle de mudanÃ§as:
- atualizaÃ§Ã£o da identificaÃ§Ã£o de versÃ£o da aplicaÃ§Ã£o para refletir a nova entrega;
- sincronizaÃ§Ã£o da versÃ£o exibida na interface, version.json e service worker;
- manutenÃ§Ã£o dos ajustes recentes de comentÃ¡rios mÃºltiplos/imÃºtaveis por atividade e correÃ§Ãµes do fluxo de salvamento.

Objetivo:
- garantir rastreabilidade adequada da liberaÃ§Ã£o no contexto de controle de mudanÃ§as;
- evitar divergÃªncia entre a versÃ£o percebida pelo usuÃ¡rio e a versÃ£o efetivamente distribuÃ­da.

## v1.2.8.8

Planejador de Recursos v1.2.8.8

Melhorias:
- Campo de comentÃ¡rios de atividade limitado a 2.000 caracteres, com contador visual.
- Tratamento de QuotaExceededError/localStorage cheio com aviso ao usuÃ¡rio.
- Limite de 10 baselines persistidos; ao atingir o limite, o sistema exige excluir um baseline antes de criar outro.
- Ajustes de persistÃªncia para reduzir falhas silenciosas ao salvar em armazenamento local.

## v1.2.8.11

Planejador de Recursos - v1.2.8.11

Principais ajustes:
- Comportamento de leitura/gravaÃ§Ã£o do banco restaurado para o padrÃ£o funcional da v1.2.8.8.
- InclusÃ£o de tolerÃ¢ncia para menos na verificaÃ§Ã£o de disponibilidade por capacidade.
- SeparaÃ§Ã£o entre recursos que atendem integralmente e recursos prÃ³ximos da meta.
- InclusÃ£o da opÃ§Ã£o Mostrar inativos na listagem de recursos.
- InclusÃ£o da aÃ§Ã£o de reativar recurso inativo diretamente pela interface.

ObservaÃ§Ã£o:
- A lÃ³gica do banco foi alinhada ao comportamento da v1.2.8.8 conforme solicitado.

## v1.2.8.12

v1.2.8.12
- ComentÃ¡rios em tabela lÃ³gica separada no BD Ãºnico.
- Compatibilidade retroativa com comentarios/comentariosJson.
- Preview leve nos cards + paginaÃ§Ã£o no modal.
- Merge por commentId na persistÃªncia CSV.

## v1.2.8.13

Planejador de Recursos - v1.2.8.13

- ComentÃ¡rios evoluÃ­dos para coleÃ§Ã£o lÃ³gica separada com persistÃªncia em comentÃ¡rio no BD Ãºnico
- Compatibilidade retroativa com comentarios e comentariosJson das atividades
- PaginaÃ§Ã£o de comentÃ¡rios no modal com render leve nos cards
- ExportaÃ§Ã£o CSV/HTML e backup/restauraÃ§Ã£o atualizados para incluir comentÃ¡rios
- ReconciliaÃ§Ã£o de comentÃ¡rios com BD apontado por commentId e preservaÃ§Ã£o de rascunho em conflito
- Reabertura da atividade com dados persistidos quando houver conflito nÃ£o resolvido

## v1.2.8.14

Planejador de Recursos - v1.2.8.14

CorreÃ§Ã£o dos comentÃ¡rios escalÃ¡veis:
- a leitura agora prioriza a tabela lÃ³gica de comentÃ¡rios quando ela jÃ¡ existe no BD
- comentÃ¡rios legados de atividade (comentarios/comentariosJson) passam a ser usados apenas como fallback de compatibilidade
- elimina duplicaÃ§Ãµes e repetiÃ§Ãµes ao reabrir atividade com dados legados + comentÃ¡rios novos no mesmo banco

## v1.2.8.15

v1.2.8.15
- Adicionada aba CalendÃ¡rio para cadastro de feriados.
- Feriados cadastrados passam a ser destacados no cabeÃ§alho e no fundo do Gantt.
- Finais de semana e feriados passam a ser considerados como dias nÃ£o Ãºteis na capacidade agregada.
- Tooltip do Gantt agora informa o tipo do dia: dia Ãºtil, final de semana ou feriado.
- Mantida retrocompatibilidade com BD existente: feriados continuam usando a estrutura date/legend jÃ¡ suportada pelo import/export.
- Atualizado service worker para forÃ§ar nova identidade de cache PWA.

## v1.2.8.16

v1.2.8.16
- Move o cadastro de feriados para uma Ã¡rea administrativa restrita dentro da aba Banco de Dados.
- Adiciona autenticaÃ§Ã£o por senha local para liberar feriados e horas extras.
- Adiciona cadastro de horas extras por recurso/data/status/justificativa.
- Ajusta a capacidade agregada para somar horas extras aprovadas Ã  capacidade diÃ¡ria real.
- MantÃ©m feriados/finais de semana como 0h base, permitindo capacidade extraordinÃ¡ria quando houver hora extra aprovada.
- Persiste horas extras no BD Ãºnico CSV/HTML usando tabela hora_extra/HorasExtras.

## v1.2.8.17

v1.2.8.17
- Removido controle de senha da seÃ§Ã£o Feriados e Horas Extras.
- Mantida a seÃ§Ã£o na aba Banco de Dados, com acesso direto.
- ReforÃ§ada a persistÃªncia de feriados e horas extras no BD apontado compartilhado.
- Corrigida exportaÃ§Ã£o HTML do BD para incluir tabela HorasExtras com cabeÃ§alhos e linhas.
- Capacidade agregada continua considerando somente horas extras aprovadas.

## v1.2.8.18

v1.2.8.18
- Corrige exclusÃ£o de lanÃ§amentos de horas extras com remoÃ§Ã£o por id e assinatura do lanÃ§amento.
- ForÃ§a salvamento no BD apontado apÃ³s remover hora extra, quando houver BD selecionado.
- Ajusta cÃ¡lculo de capacidade agregada para ignorar atividades em finais de semana/feriados sem hora extra aprovada.
- MantÃ©m capacidade adicional apenas por horas extras aprovadas.

## v1.2.8.19

v1.2.8.19
- Ajustada a visÃ£o de Carga horÃ¡ria diÃ¡ria para mostrar capacidade Ãºtil do dia, nÃ£o a alocaÃ§Ã£o bruta.
- Feriados com horas extras aprovadas passam a ter cor prÃ³pria no grÃ¡fico de capacidade agregada.
- Finais de semana com horas extras aprovadas passam a ter cor prÃ³pria no grÃ¡fico de capacidade agregada.
- AlocaÃ§Ã£o acima da capacidade aprovada Ã© indicada por linha vermelha, preservando o valor de HE como capacidade real.
- Mantida a regra: atividades em feriados/finais de semana sem HE aprovada nÃ£o entram na capacidade agregada.

