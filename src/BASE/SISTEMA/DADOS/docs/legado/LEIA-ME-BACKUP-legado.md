# Backup do banco — agendamento

Duas tarefas no **Agendador de Tarefas** de uma estação que fique ligada, ou do servidor.
O script está em `01 - PADROES\00 - IA\MIGRACAO-CRM\20 - SCRIPTS\backup-banco.ps1`.

## Agendamento — 09:00 e 16:40

Rode o **`AGENDAR-BACKUPS.bat` como administrador** (botão direito › Executar como administrador). Ele cria as duas tarefas diárias apontando para o banco de produção e, no fim, **dispara uma delas para provar que funciona** — se nenhuma cópia nova aparecer, ele avisa na hora, em vez de você descobrir semanas depois.

| Tarefa | Horário |
|---|---|
| `CRM Zapromaq - backup 0900` | 09:00, diário |
| `CRM Zapromaq - backup 1640` | 16:40, diário |

As duas usam `-Modo Periodico`: copiam mesmo com gente no sistema, registrando no log quando havia usuário conectado. Retenção de 30 dias.

### Quem executa a tarefa

Ela roda **com o seu usuário, nesta estação**, não como SYSTEM. A conta SYSTEM acessa a rede como conta de máquina e normalmente não enxerga o compartilhamento — o backup falharia todo dia, em silêncio. Como contrapartida, a tarefa depende de você estar conectado na estação.

Está marcada para **iniciar assim que possível após perder o horário**: se a máquina estiver desligada às 09:00, ela roda quando a máquina ligar.

### A limitação que continua existindo

Agendar numa estação significa depender daquela estação. Máquina desligada o dia todo, ou de férias, é um dia sem backup — e ninguém percebe, porque a ausência é silenciosa.

Três caminhos, do mais para o menos robusto:

1. **Agendar no servidor.** É onde isso deveria morar. A tarefa fica junto do arquivo e não depende de ninguém.
2. **Agendar em duas estações.** O nome do arquivo de backup leva o nome da máquina, então as cópias não se sobrescrevem; se uma estiver fora, a outra cobre. Custa duas cópias por horário.
3. **Uma estação só**, como está agora, conferindo o `backup.log` de vez em quando.

Enquanto for a opção 3, vale abrir o `backup.log` na segunda-feira e olhar se a semana toda está lá.

## Restauração — trimestral, e obrigatória antes da virada

Rode `TESTAR-RESTAURACAO.bat`. Ele pega o backup mais recente, restaura numa **cópia de teste** (nunca sobrescreve a produção), abre o banco e conta o que voltou: registros por tabela, soma de valor, quantos têm contexto e se as relações estão íntegras.

Compare com o que você espera. Se bater, o backup serve. Se o banco nem abrir, o script diz isso na cara — e aí o backup não serve, o que é melhor descobrir num teste do que no dia em que precisar.

Apague a cópia de teste depois de conferir: ela não é backup nem produção, e deixá-la lá só gera confusão sobre qual arquivo é o bom.

## O que o backup não resolve

Compactação. O `.accdb` cresce com o uso e precisa de **Compactar e Reparar** mensal, com todos fora do sistema. Backup não substitui isso.

E cópia não é histórico: se um registro for apagado por engano e ninguém perceber por uma semana, o backup horário já passou e o diário pode ter rodado sete vezes por cima. Quem protege contra isso é a tabela `log_alteracoes` e a regra de **inativar em vez de excluir**.
