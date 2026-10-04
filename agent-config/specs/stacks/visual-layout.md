# Stack: visual — layout em JSON, prévia HTML no Edge, fundos BMP

## Papel

"Desenhar em HTML, executar em MSForms": a posição e a aparência de cada tela vêm de dados
(`fonte\layout`), a prévia roda no Edge e o build embute a decoração como imagem no formulário.

## Versões e ambiente

Microsoft Edge (Chromium 132+, `--headless` novo: `--screenshot`/`--dump-dom`), PowerShell 5.1.
Sem Node: a prévia é JavaScript no próprio Edge, com os dados carregados como `.js` (`fetch` não
funciona em `file://`). Fonte de ícones: **Segoe MDL2 Assets** (Windows 10/11).

## Onde fica

```
src/BASE/SISTEMA/DADOS/fonte/layout/
  tema.json            cores, fontes, variantes de botão, cores de situação, moldura
  icones.json          glifos Segoe MDL2 (nome -> [código, nome oficial])
  amostras.json        dados FICTÍCIOS da prévia
  abas.json            abas, células (inclui Config!B2 e Início!B7)
  formularios/<tela>.json   v2: canvas (pt), fundo (decoração -> BMP), controles, componentes
src/BASE/SISTEMA/DADOS/build/gerar-visual.ps1 + visual/previa.{html,js,css}
src/BASE/SISTEMA/DADOS/fonte/assets/gerado/   GERADO: <tela>-fundo.png + manifesto.json (hash das entradas)
src/BASE/SISTEMA/DADOS/fonte/modulos/modTema.bas   GERADO
```

## Comandos

```powershell
powershell -File src\BASE\SISTEMA\DADOS\build\gerar-visual.ps1 [-Tela frmCRM] [-Tolerante]
# prévias: src\BASE\SISTEMA\DADOS\execucao\previa\<tela>-100|125|150.png e -contornos.png
```

O build roda o `gerar-visual` sozinho se o manifesto estiver velho, converte PNG → BMP 96 dpi e
embute no `Designer.Picture` em tempo de desenho.

## Estilo

- Projeto a 100%; 1 pt ≈ 1,333 px a 96 dpi; tudo cabe em 1366×768 (o `modTela.Encaixar` ajusta o zoom).
- Texto e ícones são nativos (Labels transparentes + glifos MDL2); só a decoração (cartões,
  sombras, degradês) vira imagem.
- Nome de controle no JSON é contrato com `formularios/<tela>.txt` (o `verificar.py` confere).
- Gerados não se editam: mude o JSON/tema e rode o `gerar-visual`.

## Testes

A conferência do `gerar-visual` (sobreposição, fora da área, nomes) roda em toda geração e
falha o build. Aparência real: comparar a prévia com `frmCalibracao` numa estação 🖥 (100% e 125%).

## Limites

- Nunca: `LoadPicture` de disco em execução (o `.xlsm` roda de `Documentos`); imagem que não
  seja bmp/gif/jpg no `Picture`; editar `assets/gerado` ou `modTema.bas` à mão; dado real em `amostras.json`.

## Revisão

- [ ] Manifesto regenerado e commitado junto com a mudança de layout.
- [ ] Prévias a 100/125/150% conferidas; nada cortado em 1366×768.
- [ ] Controles novos têm o mesmo nome no JSON e no código do formulário.

## Referências offline

`referencias/` 10-edge-headless, 11-windows-icones-dpi, 01-vba-msforms.
