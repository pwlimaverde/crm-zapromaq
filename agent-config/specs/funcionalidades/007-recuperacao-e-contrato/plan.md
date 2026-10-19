# Plano: Recuperação, manutenção e contrato do conector (ciclo 007)

Spec: [SPEC.md](SPEC.md)

## Ordem e dependências

1. DEV-07 contrato exato e capacidades no `crm_status` — depende de: —
2. DEV-08 guarda de manutenção no conector e no front; montagem liga o sinal — depende de: 1
3. DEV-03 exclusividade durante a cópia + mensagem controlada (conector e front) — depende de: 2
4. DEV-03 janela atendida + aviso por janela (PS, conector, front) — depende de: 3
5. DEV-06A diagnóstico do front, retenção e falha no `crm_status` — depende de: 3
6. DEV-04 volta e manifesto de release — depende de: —
7. DEV-05 conector por pasta de versão e ponteiro — depende de: 1
8. Documentação, nota de mudança, PLANO (D24, D25, D23) — depende de: 1-7

## Checkpoints

- Depois da 3: experimento de exclusividade com ACE real no `testar-recuperacao`.
- Depois da 8: `verificar-tudo -Completo`.
- 🖥 publicação 1.6, restauração cronometrada, volta do front e do conector na estação.

## Riscos

- O ACE abrir só leitura com o arquivo seguro muda de comportamento em outra versão do Office: o teste com ACE real falha e a mensagem controlada deixa de aparecer.
- Teste de quatro processos simultâneos ser lento: espera curta por parâmetro.
