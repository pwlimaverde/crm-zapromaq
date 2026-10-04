# Implantação na rede — roteiro (Fase 8)

Roteiro para trocar a versão 1.4 (legado, `MIGRACAO-CRM`) pela estrutura nova.
Feito **fora do horário de uso**, com todos os vendedores com o CRM fechado.
Tempo estimado: 1 h, mais a ida às estações. Cada passo tem a conferência que
libera o seguinte — não pule conferência.

Legenda: 🖥 = numa estação com Excel 2019 · 🌐 = na pasta da rede

## 0. Antes do dia

- [ ] Idas à estação nº 1 e nº 2 feitas (PLANO 2.10 e 4.11): build conferido e
      calibração visual aprovada.
- [ ] Definido onde fica o `INICIAR-CRM.bat` das estações (pendência P01): ele
      **não pode** ficar na raiz de `BASE` (só `crm_zapromaq.accdb`,
      `CRM_Zapromaq.xlsm`, `SISTEMA\` e `IA\`). Sugestão: a pasta onde já está hoje,
      apontando para o novo caminho de `BASE\CRM_Zapromaq.xlsm`.
- [ ] Repositório sincronizado com a rede (`docs\sincronizar-rede.md`) e com
      `verificar.py` sem problemas.

## 1. Backup da produção atual (8.1) 🌐

1. Copiar a pasta de produção inteira (legado) para
   `...\BACKUP-ANTES-DA-V1.5-AAAAMMDD\` — `.accdb`, `.xlsm`, fontes e scripts.
2. Conferir: o `.accdb` copiado abre (tamanho igual ao original) e não existe
   `.laccdb` ao lado do original (ninguém com o banco aberto).

## 2. Pasta nova na rede (8.2) 🌐

1. Copiar a pasta `BASE` do repositório para a rede **sem** `crm_zapromaq.accdb`
   e **sem** `CRM_Zapromaq.xlsm` (não existem no repositório) e sem
   `SISTEMA\DADOS\execucao\` (é gerada na rede).
2. **Mover** (não copiar) o `.accdb` de produção para a raiz da `BASE` nova, com o
   nome `crm_zapromaq.accdb`. Mover garante que ninguém continue gravando no antigo.
3. 🖥 `SISTEMA\DADOS\testes\CHECAR-AMBIENTE.bat` a partir da rede: provedor
   ACE, Excel 64 bits, área útil da tela.
4. 🖥 `SISTEMA\DADOS\banco\APLICAR-MIGRACOES.bat`: aplica as migrações que faltam
   (a 002 reativa os cadastros criados inativos pelo defeito A01). Conferir no
   log que `config.versao_esquema` ficou na versão esperada.

## 3. Front na rede (8.3) 🖥

1. `SISTEMA\MONTAR-FRONTEND.bat -Teste` → abrir
   `SISTEMA\DADOS\execucao\teste\CRM_Zapromaq.xlsm`: Alt+F8 › `AbrirCalibracao`
   (confere o visual) e uma consulta em cada aba.
   Antes disso, o teste automático faz o mesmo sem abrir a tela:
   `powershell -File SISTEMA\DADOS\testes\testar-uso.ps1` (grades, Painel,
   "Ver tudo" com filtro, relatório e PDFs em `execucao\teste\uso-*.pdf`).
   Se a montagem falhar na compilação:
   `powershell -File SISTEMA\DADOS\testes\diagnosticar-compilacao.ps1` diz o
   módulo, a linha e a mensagem.
2. `SISTEMA\MONTAR-FRONTEND.bat` (publicação): monta, compila, roda o autoteste
   e só então grava `BASE\CRM_Zapromaq.xlsm` e sobe a versão para **1.5**.
3. Conferir `Início!B7` do arquivo publicado:
   `Ambiente: PRODUCAO   |   Versao 1.5   |   Publicada em dd/MM/yyyy HH:mm`.
4. Atualizar o `INICIAR-CRM.bat` das estações para o caminho novo
   (`\\servidor\...\BASE\CRM_Zapromaq.xlsm`).
5. Em **cada** estação: abrir pelo atalho de sempre. Conferir que copiou a 1.5
   para `Documentos\CRM Zapromaq\`, que as grades carregam e que a ficha abre.
6. Com duas estações abertas, gravar numa e ver a linha mudar na outra em até
   30 s (PLANO 6.5).
7. Na aba Painel, **Exportar relatório**: confere que a pasta sai em
   `Documentos\CRM Zapromaq\Relatorios\<data-hora>\` com os CSVs, o
   `painel.pdf` e o `LEIA-ME.txt`.

## 4. IA (8.4) 🖥

Em cada máquina com Claude Desktop, **a partir da pasta da rede**:

1. `IA\INSTALAR-MCP.bat` (grava o caminho UNC na configuração do Claude Desktop).
2. Fechar o Claude Desktop por completo e abrir; conferir em **+ › Conectores**
   `crm-zapromaq` com 11 ferramentas; pedir *"situação do banco do CRM"*.
3. Instalar a skill (`IA\LEIA-ME.md`, passo 6).
4. Só na máquina de quem mantém o sistema: `IA\INSTALAR-MCP-DEV.bat`
   (`crm-zapromaq-dev`, 15 ferramentas).

## 5. Backups (8.5 e 8.6) 🖥

1. Remover o agendamento antigo (Agendador de Tarefas do usuário: tarefas do
   legado, que apontam para o caminho velho).
2. `SISTEMA\DADOS\operacao\AGENDAR-BACKUPS.bat` a partir da rede (sem
   administrador; roda com o usuário logado, 09:00 e 16:40 por padrão).
3. `SISTEMA\DADOS\operacao\FAZER-BACKUP.bat` uma vez agora; conferir o arquivo em
   `SISTEMA\DADOS\execucao\backup\`.
4. `SISTEMA\DADOS\operacao\TESTAR-RESTAURACAO.bat`: restaura o último backup numa
   cópia e confere que abre e tem as tabelas. **Backup não testado não é backup.**

## Voltar atrás (se algo der errado)

Até o passo 3.4 as estações continuam usando a 1.4: basta mover o `.accdb` de volta
para a pasta antiga. Depois do 3.4: devolver o `.accdb` à pasta antiga, restaurar o
`INICIAR-CRM.bat` anterior e pedir aos vendedores que reabram. Dados gravados na
1.5 continuam no banco (o esquema novo só acrescenta, não remove).

## Depois

- Acompanhar os logs em `SISTEMA\DADOS\execucao\logs\` e `IA\logs\` na primeira
  semana.
- Ajustes pequenos: MCP de desenvolvimento na rede; mudanças maiores: repositório.
  Em ambos os casos, a regra de sincronização (`docs\sincronizar-rede.md`).
