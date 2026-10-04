# SISTEMADADOS — front do CRM Zapromaq (fonte, build e operação)

Tudo que monta e mantém o `CRM_Zapromaq.xlsm` e o banco `crm_zapromaq.accdb` da raiz
de `BASE`. Roda só com o que existe numa estação: **Windows PowerShell 5.1 + Excel 2019
64 bits + Edge**. Nenhum caminho absoluto: copiar `BASE` inteira para a rede basta.

```
SISTEMA\
  MONTAR-FRONTEND.bat          monta o .xlsm da raiz e sobe a versão (1.4 -> 1.5)
  DADOS\
    VERSAO.txt                 versão publicada do front (só o build altera)
    fonte\                     o que vira o .xlsm (VBA em Windows-1252 + CRLF)
      modulos\  classes\  formularios\  pasta-de-trabalho\
      layout\                  tema e posição dos elementos de cada tela
      assets\                  logo e imagens; assets\gerado\ = fundos gerados
    build\                     montar-frontend.ps1, gerador visual
    lib\Comum.ps1              funções compartilhadas (caminhos, UNC, ACE, log)
    banco\
      esquema\criar-banco.ps1  estrutura do banco (esquema 1.0 = produção)
      migracoes\NNN-*.ps1      mudanças posteriores, em ordem
      APLICAR-MIGRACOES.bat    aplica as pendentes (com backup antes)
    operacao\
      FAZER-BACKUP.bat         backup manual (completo ou simples)
      TESTAR-RESTAURACAO.bat   restaura o último backup numa cópia e confere
      AGENDAR-BACKUPS.bat      tarefas 09:00, 12:30 e 16:00, seg-sex, para o usuário (sem administrador);
                               retenção: mês corrente e anterior
    testes\
      CHECAR-AMBIENTE.bat      o que esta máquina tem e o que falta
      conferir-painel.ps1      indicadores do Painel × banco
    ferramentas\               só desenvolvimento (Python): verificação, registro da migração 3.6
    docs\                      contrato do banco, CHANGELOG, documentos do legado
    execucao\                  gerado (fora do git): backup\ historico-front\ logs\ previa\
```

## Regras

- A raiz de `BASE` tem **só** `crm_zapromaq.accdb`, `CRM_Zapromaq.xlsm`, `SISTEMA\` e `IA\`.
- O `.xlsm` tem nome fixo. A célula `Início!B7` mantém o texto exato
  `Ambiente: PRODUCAO   |   Versao X.Y   |   Publicada em dd/MM/yyyy HH:mm` —
  o `.bat` das estações compara esse texto para saber se copia a versão nova.
- **Fonte é texto; `.xlsm` e `.accdb` são produto.** Nunca editar o `.xlsm` à mão.
- Regras de gravação no banco: `DADOS\docs\contrato-banco.md`.
