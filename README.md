# Simulador de Investimentos em Fundos Imobiliários (FIIs)

Projeto do desafio "Simulador de Investimentos em Fundos Imobiliários" da trilha de Excel na [DIO](https://www.dio.me/). O objetivo é construir uma planilha que ajude a simular aportes mensais em FIIs, projetar patrimônio acumulado e dividendos ao longo do tempo, e sugerir uma alocação de carteira conforme o perfil de risco do investidor.

## O que a planilha faz

A `Planilhia_de_FII.xlsx` tem duas abas:

### Página1 — Simulador

**Configurações**
- Salário mensal e percentual destinado a aportes (30% do salário, calculado automaticamente).

**Investimento Mensal**
- Valor do aporte mensal, prazo em anos e rendimento mensal (taxa) informados manualmente.
- Patrimônio acumulado calculado com `=FV(taxa, período, -aporte)` — a função financeira de valor futuro do Excel.
- Dividendo mensal estimado = patrimônio acumulado × rendimento (dividend yield).

**Cenários**
- Projeção do patrimônio acumulado e dos dividendos para 2, 5, 10, 15, 20 e 30 anos de aporte, todos usando a mesma taxa e o mesmo aporte mensal — só varia o prazo.

**Perfil de carteira**
- Escolha de perfil (Conservador, Moderado ou Agressivo).
- `VLOOKUP` busca na Página2 o percentual sugerido de cada tipo de FII (Papel, Tijolo, Híbrido, FOFs, Desenvolvimento, Hotelaria) para o perfil escolhido.
- Multiplica o percentual pelo valor mensal disponível para aporte, retornando quanto alocar em cada tipo de fundo.

### Página2 — Tabela de referência

Matriz perfil × tipo de FII com o percentual sugerido de alocação para cada combinação (linha auxiliar `Chave` concatena perfil e tipo para o `VLOOKUP` funcionar). É aqui que fica a lógica de "quanto de cada tipo de fundo para cada perfil de risco".

## Conceitos de Excel aplicados

- Função financeira `FV` (valor futuro) para projeção de patrimônio com aportes constantes
- `VLOOKUP` com chave concatenada para simular uma busca por múltiplos critérios
- Referências absolutas (`$E$12`, `$A:$D`) para fórmulas replicáveis
- Separação de dados de entrada, cálculos e tabela de apoio em abas distintas

## Como usar

1. Baixe o arquivo `Planilhia_de_FII.xlsx`.
2. Na aba **Página1**, preencha salário, aporte mensal, prazo e rendimento mensal esperado.
3. Escolha o perfil de investidor (Conservador, Moderado, Agressivo).
4. Veja o patrimônio projetado, os dividendos estimados e a sugestão de alocação por tipo de FII.

## Estrutura do repositório

```
.
├── README.md
├── Planilhia_de_FII.xlsx
└── images/
    └── (capturas de tela, se houver)
```

## Observação sobre os perfis de risco

A classificação Conservador/Moderado/Agressivo usada é uma simplificação didática do desafio, não uma recomendação real de alocação. Na prática, FIIs não são um ativo "conservador" — mesmo um fundo de papel (recebíveis) carrega risco de crédito e de marcação a mercado, e a rotulagem por perfil aqui serve só para exercitar o `VLOOKUP` com chave composta, não reflete uma análise de risco de fato.

## Aprendizados

Este desafio fixou o uso de `FV` para projeção de investimentos recorrentes e de `VLOOKUP` com chave composta para tabelas de decisão (perfil × alocação) — uma alternativa simples a `INDEX/MATCH` quando se quer buscar por mais de uma coluna.
