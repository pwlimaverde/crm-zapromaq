# Distribuição do CRM para as estações

Versão 1.4 · 18/09/2026

## O caminho normal: o atalho CRM Zapromaq

`INICIAR-CRM.bat` **fica só na pasta da rede**, em `01 - CRM\01 - CONTROLE\BASE`, junto do
modelo e do `crm_zapromaq.ico`. Nenhuma estação recebe cópia dele — o que vai para a
máquina é um **atalho** que aponta para ele na rede. Assim, corrigir o lançador é mexer
num arquivo só, e não em cada estação.

Na primeira vez, a pessoa abre o `INICIAR-CRM.bat` direto da pasta da rede. Ele cria o
atalho **CRM Zapromaq** em `Documentos\CRM Zapromaq\`, com o logo da empresa como ícone,
e a partir daí é só usar o atalho — que pode ser arrastado para a área de trabalho.

A cada execução ele:

1. acha a pasta da rede — tenta, nesta ordem, a própria pasta onde está, a unidade
   mapeada e o caminho UNC; a primeira que existir vence, então não há letra de
   unidade fixa no código;
2. acha a pasta **Documentos deste usuário** pelo registro do Windows — não por
   `%USERPROFILE%\Documents`, porque com OneDrive ligado a pasta é redirecionada e o
   caminho montado à mão aponta para o lugar errado;
3. cria o atalho, se ainda não existir, e copia o ícone para a máquina — atalho com
   ícone em caminho de rede fica genérico quando a rede não responde;
4. copia o `CRM_Zapromaq.xlsm` da rede **só se a versão de lá for diferente**;
5. abre a cópia local.

A cópia de trabalho fica em `Documentos\CRM Zapromaq\`.

## O que acontece quando algo falha

| Situação | Comportamento |
|---|---|
| Rede fora do ar | Abre a cópia local que existir, avisando que pode estar desatualizada |
| CRM já aberto na estação | Não consegue substituir o arquivo; avisa para fechar o Excel e abre a cópia atual |
| Nunca houve cópia e a rede não responde | Não abre nada e explica o motivo |

## Por que não é o usuário que copia da rede

O modelo da rede **não** é para trabalhar: duas pessoas no mesmo arquivo foi o que
corrompeu a planilha compartilhada. Quem abre o modelo direto recebe o aviso e a oferta
de instalar a cópia local; o salvamento fica bloqueado.

Além disso, cada montagem grava a versão publicada na tabela `config` do banco. A cópia
da estação compara na abertura e avisa quando ficou para trás — o aviso agora manda usar
o atalho INICIAR-CRM, não ir à pasta da rede.

## Quando publicar uma versão nova

1. `MONTAR-FRONTEND.bat` com o `.xlsm` fechado em todas as estações.
2. Nada mais. Cada estação se atualiza no próximo INICIAR-CRM.

Não é preciso avisar a equipe nem copiar arquivo em máquina nenhuma.
