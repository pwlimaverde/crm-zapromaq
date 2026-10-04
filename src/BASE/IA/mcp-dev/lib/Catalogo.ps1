<#
  Catalogo.ps1 - lista FECHADA de ferramentas do MCP de desenvolvimento.
  Cada handler está em Fontes.ps1 ou Migracoes.ps1.
#>

$script:Ferramentas = @(
    [ordered]@{
        name = 'listar_fontes'; title = 'Listar fontes do front'; handler = 'Invoke-ListarFontes'; annotations = $script:SoLeitura
        description = 'Lista os arquivos editáveis de SISTEMA\DADOS\fonte (VBA .bas/.cls/.txt, layout .json). Use "pasta" para filtrar (ex.: layout\formularios).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            pasta = [ordered]@{ type = 'string'; maxLength = 120; description = 'Subpasta de fonte\, ex.: modulos' } } }
    },
    [ordered]@{
        name = 'ler_fonte'; title = 'Ler um arquivo de fonte'; handler = 'Invoke-LerFonte'; annotations = $script:SoLeitura
        description = 'Devolve o conteúdo e a "versao" (hash) de um arquivo de fonte\. A versao é obrigatória para gravar_fonte.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('caminho'); properties = [ordered]@{
            caminho = [ordered]@{ type = 'string'; maxLength = 200; description = 'Relativo a fonte\, ex.: layout\formularios\frmCRM.json' } } }
    },
    [ordered]@{
        name = 'gravar_fonte'; title = 'Gravar um arquivo de fonte'; handler = 'Invoke-GravarFonte'; annotations = $script:Grava
        description = 'Substitui o conteúdo INTEIRO de um arquivo de fonte\ (ou cria um novo, sem "versao"). Exige a versao lida em ler_fonte; guarda cópia do anterior. VBA é gravado em Windows-1252 + CRLF. Arquivos gerados (assets\gerado, modTema.bas) são recusados. Mostre a alteração ao usuário antes.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('caminho', 'conteudo'); properties = [ordered]@{
            caminho = [ordered]@{ type = 'string'; maxLength = 200 }
            conteudo = [ordered]@{ type = 'string'; maxLength = 400000 }
            versao = [ordered]@{ type = 'string'; maxLength = 16; description = 'Versão devolvida por ler_fonte (omitir só para arquivo novo)' }
            motivo = [ordered]@{ type = 'string'; maxLength = 300; description = 'O que foi ajustado e por quê (vai para o registro de alterações)' } } }
    },
    [ordered]@{
        name = 'gerar_previa'; title = 'Prévia visual de uma tela'; handler = 'Invoke-GerarPrevia'; annotations = $script:SoLeitura
        description = 'Gera a prévia da tela a partir de layout\ (Edge headless) e devolve a IMAGEM, mais o relatório de problemas (texto que não cabe, sobreposição, fora da tela). Também regera fundos e modTema.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tela'); properties = [ordered]@{
            tela = [ordered]@{ type = 'string'; pattern = '^frm[A-Za-z0-9]+$'; description = 'Nome do formulário, ex.: frmCRM' }
            escala = [ordered]@{ type = 'integer'; enum = @(100, 125, 150); default = 100 } } }
    },
    [ordered]@{
        name = 'verificar_layout'; title = 'Conferir layout'; handler = 'Invoke-VerificarLayout'; annotations = $script:SoLeitura
        description = 'Roda o gerador visual em todas as telas (ou em uma) e devolve só a conferência: sem_problemas e a lista do que não cabe ou se sobrepõe.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            tela = [ordered]@{ type = 'string'; pattern = '^frm[A-Za-z0-9]+$' } } }
    },
    [ordered]@{
        name = 'montar_teste'; title = 'Montar .xlsm de teste'; handler = 'Invoke-MontarTeste'; annotations = $script:Grava
        description = 'Monta o front em SISTEMA\DADOS\execucao\teste\ (compila e roda o autoteste) SEM publicar nem subir a versão. Exige Excel nesta máquina. Leva de 1 a 3 minutos.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'publicar'; title = 'Publicar o front para todos'; handler = 'Invoke-Publicar'; annotations = $script:Publica
        description = 'Monta e PUBLICA o CRM_Zapromaq.xlsm na raiz (sobe a versão; as estações copiam na próxima abertura). Só com pedido explícito do usuário, depois de montar_teste dar certo. Exige confirmacao = "PUBLICAR".'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('confirmacao'); properties = [ordered]@{
            confirmacao = [ordered]@{ type = 'string'; enum = @('PUBLICAR') } } }
    },
    [ordered]@{
        name = 'listar_alteracoes'; title = 'Alterações feitas por aqui'; handler = 'Invoke-ListarAlteracoes'; annotations = $script:SoLeitura
        description = 'Registro das gravações feitas por este servidor (quando, quem, arquivo, motivo). É a lista do que precisa voltar ao repositório de desenvolvimento.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            limite = [ordered]@{ type = 'integer'; minimum = 1; maximum = 500; default = 50 } } }
    },
    [ordered]@{
        name = 'estado_sistema'; title = 'Estado do sistema'; handler = 'Invoke-EstadoSistema'; annotations = $script:SoLeitura
        description = 'Versões (próxima publicação, publicada na raiz, no banco), versão do esquema, ambiente, migrações pendentes e aplicadas, se o banco está em uso e se há Excel na máquina. Use antes de qualquer alteração.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'listar_migracoes'; title = 'Migrações do banco'; handler = 'Invoke-ListarMigracoes'; annotations = $script:SoLeitura
        description = 'Arquivos de banco\migracoes com id, descrição e se já foram aplicados neste banco.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'ler_migracao'; title = 'Ler uma migração'; handler = 'Invoke-LerMigracao'; annotations = $script:SoLeitura
        description = 'Conteúdo de um arquivo de migração (para conferir antes de aplicar).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('arquivo'); properties = [ordered]@{
            arquivo = [ordered]@{ type = 'string'; maxLength = 120; description = 'Nome do arquivo, ex.: 003-campo-novo.ps1' } } }
    },
    [ordered]@{
        name = 'criar_migracao'; title = 'Criar migração do banco'; handler = 'Invoke-CriarMigracao'; annotations = $script:Grava
        description = 'Cria o próximo arquivo numerado em banco\migracoes a partir dos comandos SQL informados. tipo=dados (roda em transação, pode ser aplicada com gente usando) ou tipo=estrutura (campo/tabela nova: exige banco exclusivo, esquema_para e publicação de front e IA na mesma versão). Alterar a estrutura do banco à mão é proibido: é sempre por migração. DROP, DELETE ou UPDATE sem WHERE exigem confirmacao = "CONFIRMO".'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('nome', 'descricao', 'sql'); properties = [ordered]@{
            nome = [ordered]@{ type = 'string'; maxLength = 60; description = 'Parte do nome do arquivo, ex.: campo-cnae-em-clientes' }
            descricao = [ordered]@{ type = 'string'; maxLength = 200; description = 'O que a migração faz (aparece no log)' }
            tipo = [ordered]@{ type = 'string'; enum = @('dados', 'estrutura'); default = 'dados' }
            esquema_para = [ordered]@{ type = 'string'; maxLength = 10; description = 'Nova versão do esquema (só em tipo=estrutura), ex.: 1.1' }
            sql = [ordered]@{ type = 'array'; description = 'Comandos SQL na ordem (ACE/Access)'; items = [ordered]@{ type = 'string'; maxLength = 2000 } }
            confirmacao = [ordered]@{ type = 'string'; enum = @('CONFIRMO'); description = 'Obrigatório para comandos destrutivos' } } }
    },
    [ordered]@{
        name = 'aplicar_migracoes'; title = 'Aplicar migrações pendentes'; handler = 'Invoke-AplicarMigracoes'; annotations = $script:Grava
        description = 'Aplica as migrações ainda não aplicadas (faz backup do banco antes). simular=true só mostra o que seria feito. A publicação (publicar) também aplica sozinha o que estiver pendente.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{
            simular = [ordered]@{ type = 'boolean'; default = $false } } }
    },
    [ordered]@{
        name = 'verificar_projeto'; title = 'Conferir o projeto'; handler = 'Invoke-VerificarProjeto'; annotations = $script:SoLeitura
        description = 'Conferência estática antes de montar: codificação do VBA, controles citados no código x layout, campos de modSchema x DDL x consultas, sintaxe dos .ps1, .bat e JSON. Tudo em PowerShell (sem Python). Sintaxe VBA fica com montar_teste.'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; properties = [ordered]@{} }
    },
    [ordered]@{
        name = 'ver_log'; title = 'Último log'; handler = 'Invoke-VerLog'; annotations = $script:SoLeitura
        description = 'Final do último log: build (publicação), visual (prévias), migracoes, backup ou ambiente (execucao\logs).'
        inputSchema = [ordered]@{ type = 'object'; additionalProperties = $false; required = @('tipo'); properties = [ordered]@{
            tipo = [ordered]@{ type = 'string'; enum = @('build', 'visual', 'migracoes', 'backup', 'ambiente') } } }
    }
)
