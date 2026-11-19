descricao: Passo a passo e regras técnicas para montar uma rede (postes, estruturas MT/BT, chaves, trafos, para-raios, aterramento) a partir de uma descrição do usuário.
# Montar um projeto

Pedidos típicos: "monte uma rede com 3 postes...", "adicione um trafo de 75 kVA no poste 2". Siga esta ordem e proponha tudo de uma vez com `propor_alteracoes` (o usuário confirma).

## Passo a passo
1. `ler_tabela` — veja o que já existe (não repita nem duplique).
2. Se o usuário falou em termos soltos ("poste de 11", "trafo de 25", "chave"), resolva com `buscar_ativo` e a tabela de interpretação abaixo.
3. Monte UMA linha de Outros por poste, no formato da skill `formato_ativos` (ex.: `11300D, N1, 1 CFU, 1 TR125`): poste primeiro, depois as estruturas, depois equipamentos. Ativos soltos (sem poste) ficam agrupados em linhas separadas, com a quantidade total.
4. Cabos entram na tabela Cabos (um vão por linha) com fases e comprimento informados; se faltar o comprimento, PERGUNTE.
5. `validar_linhas` nas linhas novas e corrija o que a base não reconhecer.
6. `propor_alteracoes`; depois, se pedirem valores, `calcular_orcamento` com `com_plano=true`.

## Interpretação de termos (padrões mais comuns — confirme na base)
- "Poste de 11 metros" → `DT11/300`. "Trafo de 25" → `TR125` (mono 25 kVA). "Chave" → `CFU`. "Elo de 3H" → `EF3H`.
- "Para-raio de MT" → `PR15`; de BT → `PR220`. "Aterramento" sem detalhe → `TERRA3`.

## Regras técnicas
- Postes em MT: nunca 10 m ou menos; use DT11/300 ou superior.
- Estruturas MT por número: 1 e 2 passante; 3 fim de rede/derivação; 4 amarração. Estruturas de fim (U3, N3, R3…) não ficam sozinhas no poste: vêm com outra estrutura MT ou um transformador.
- Extensão BT típica: poste 1 `SI4, SI3, RA2`; poste 2 `SI3, RA2` (no compacto: `11300D, SI4, SI3, RA2`) (SI1/SI2 passante, SI3 fim/derivação, SI4 amarração).
- CFU monofásica: sempre 1 SUPL antes da chave, e o elo `EF…` no mesmo poste. Poste com trafo não leva elo, a menos que também tenha chave.
- Transformador leva `PR15` (1) (não use se houver reinstalação `RPR`), `2 PR220` (mono) ou `3 PR220` (tri); com duas descidas, dobre os PR220.
- P50 do trafo: base 2 m (mono) ou 6 m (tri) + 1 m por PR15.

## Reinstalação e abertura (códigos com prefixo)
- `R…` reinstala (RCFU, RCFA, RTR1…); `A…` abre (ACFU, ACFA, ATR1…). Confirme o código exato na base.

## O que NÃO fazer
- Não invente ativos, quantidades de equipamento ou comprimentos. Não aplique nada sem o plano ser confirmado.
