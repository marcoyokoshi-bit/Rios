const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, Footer, Header, PageBreak
} = require('/tmp/node_modules/docx');
const fs = require('fs');

// ─── Colors ────────────────────────────────────────────────────────────────
const C = {
  dark:    "1A1A2E",
  accent:  "E94560",
  phase1:  "16213E",
  phase2:  "0F3460",
  phase3:  "533483",
  phase4:  "2B2D42",
  white:   "FFFFFF",
  light:   "F4F6F9",
  mid:     "E2E8F0",
  text:    "2D3748",
  sub:     "718096",
  green:   "38A169",
  yellow:  "D69E2E",
  red:     "E53E3E",
};

const border0 = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const borders0 = { top: border0, bottom: border0, left: border0, right: border0 };
const borderGray = { style: BorderStyle.SINGLE, size: 1, color: C.mid };
const bordersGray = { top: borderGray, bottom: borderGray, left: borderGray, right: borderGray };

// ─── Helpers ───────────────────────────────────────────────────────────────
const run = (text, opts = {}) => new TextRun({ text, font: "Arial", ...opts });
const para = (children, opts = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
const spacer = (before = 80) => para([run("")], { spacing: { before, after: 0 } });

function sectionHeader(title, subtitle, bgColor) {
  return [
    spacer(200),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders: borders0,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 300, right: 300 },
        children: [
          para([run(title, { bold: true, size: 32, color: C.white })]),
          para([run(subtitle, { size: 22, color: C.white, italics: true })]),
        ]
      })]})],
    }),
  ];
}

function moduleCard(num, title, tagline, items, status) {
  const statusColors = { MVP: C.green, Fase2: C.accent, Fase3: C.phase3, Fase4: C.dark };
  const statusBg = statusColors[status] || C.sub;
  const itemRows = items.map(it =>
    new TableRow({ children: [
      new TableCell({
        borders: borders0,
        width: { size: 500, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 120, right: 80 },
        children: [para([run("•", { color: C.accent, size: 20 })])],
      }),
      new TableCell({
        borders: borders0,
        width: { size: 7660, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 0, right: 120 },
        children: [para([run(it, { size: 20, color: C.text })])],
      }),
    ]})
  );

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: bordersGray,
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: C.light, type: ShadingType.CLEAR },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [
        // header strip
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [700, 7560, 1100],
          rows: [new TableRow({ children: [
            new TableCell({
              borders: borders0,
              width: { size: 700, type: WidthType.DXA },
              shading: { fill: statusBg, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 120, bottom: 120, left: 0, right: 0 },
              children: [para([run(num, { bold: true, size: 28, color: C.white })], { alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              borders: borders0,
              width: { size: 7560, type: WidthType.DXA },
              shading: { fill: C.white, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 80, left: 160, right: 120 },
              children: [
                para([run(title, { bold: true, size: 24, color: C.text })]),
                para([run(tagline, { size: 19, color: C.sub, italics: true })]),
              ],
            }),
            new TableCell({
              borders: borders0,
              width: { size: 1100, type: WidthType.DXA },
              shading: { fill: C.white, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 100, bottom: 80, left: 0, right: 120 },
              children: [para([run(status, { bold: true, size: 17, color: statusBg })], { alignment: AlignmentType.RIGHT })],
            }),
          ]})]
        }),
        // body
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [500, 8860],
          rows: [
            ...itemRows,
            new TableRow({ children: [
              new TableCell({ borders: borders0, width: { size: 500, type: WidthType.DXA }, children: [para([])] }),
              new TableCell({ borders: borders0, width: { size: 8860, type: WidthType.DXA }, margins: { bottom: 80 }, children: [para([])] }),
            ]}),
          ],
        }),
      ],
    })]})],
  });
}

function phaseTable(phases) {
  const rows = phases.map(([phase, color, modules, prazo]) =>
    new TableRow({ children: [
      new TableCell({
        borders: bordersGray,
        width: { size: 2200, type: WidthType.DXA },
        shading: { fill: color, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 120 },
        children: [para([run(phase, { bold: true, size: 22, color: C.white })])],
      }),
      new TableCell({
        borders: bordersGray,
        width: { size: 5360, type: WidthType.DXA },
        margins: { top: 120, bottom: 120, left: 160, right: 120 },
        children: [para([run(modules, { size: 20, color: C.text })])],
      }),
      new TableCell({
        borders: bordersGray,
        width: { size: 1800, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        children: [para([run(prazo, { bold: true, size: 20, color: C.sub })], { alignment: AlignmentType.CENTER })],
      }),
    ]}),
  );

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2200, 5360, 1800],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: bordersGray, width: { size: 2200, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 120 }, children: [para([run("FASE", { bold: true, size: 20, color: C.white })])] }),
        new TableCell({ borders: bordersGray, width: { size: 5360, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 120 }, children: [para([run("MÓDULOS", { bold: true, size: 20, color: C.white })])] }),
        new TableCell({ borders: bordersGray, width: { size: 1800, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [para([run("PRAZO EST.", { bold: true, size: 20, color: C.white })], { alignment: AlignmentType.CENTER })] }),
      ]}),
      ...rows,
    ],
  });
}

// ─── Document ──────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: C.text } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 36, bold: true, font: "Arial", color: C.dark }, paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: C.text }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({ children: [
        para([
          run("RIOS — Restaurant Intelligence OS  |  Chef Marco À Souza  |  Confidencial", { size: 16, color: C.sub }),
          run("   ", { size: 16 }),
        ], { alignment: AlignmentType.CENTER }),
      ]}),
    },
    children: [

      // ── CAPA ──────────────────────────────────────────────────────────
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: borders0,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: C.dark, type: ShadingType.CLEAR },
          margins: { top: 800, bottom: 800, left: 500, right: 500 },
          children: [
            para([run("RIOS", { bold: true, size: 96, color: C.accent })], { alignment: AlignmentType.CENTER }),
            para([run("Restaurant Intelligence OS", { bold: true, size: 36, color: C.white })], { alignment: AlignmentType.CENTER }),
            spacer(120),
            para([run("O primeiro sistema que deixa seu restaurante mais inteligente.", { size: 26, color: C.mid, italics: true })], { alignment: AlignmentType.CENTER }),
            spacer(200),
            para([run("ROADMAP ESTRATÉGICO DE DESENVOLVIMENTO", { bold: true, size: 20, color: C.accent })], { alignment: AlignmentType.CENTER }),
            spacer(60),
            para([run("Versão 1.0  |  Chef Marco À Souza  |  Junho 2026", { size: 18, color: C.sub })], { alignment: AlignmentType.CENTER }),
          ],
        })]})],
      }),

      spacer(400),

      // ── VISÃO GERAL ────────────────────────────────────────────────────
      para([run("O QUE É O RIOS", { bold: true, size: 28, color: C.dark })]),
      para([
        run("Não é um ERP. Não é um sistema financeiro. Não é um sistema de estoque.", { bold: true, size: 22, color: C.accent }),
      ]),
      para([run("É uma inteligência que conecta todos eles — e vai além.", { size: 22, color: C.text })]),
      spacer(80),
      para([run(
        "Enquanto os ERPs respondem 'o que aconteceu', o RIOS responde três perguntas mais valiosas: " +
        "Por que aconteceu? O que vai acontecer se nada mudar? Qual é a melhor mudança possível antes que o problema apareça?",
        { size: 22, color: C.text }
      )]),

      spacer(200),

      // ── SUMÁRIO DE FASES ──────────────────────────────────────────────
      para([run("VISÃO GERAL DO ROADMAP", { bold: true, size: 26, color: C.dark })]),
      spacer(80),
      phaseTable([
        ["FASE 1\nFundação", C.green,   "Estoque · Cardápio · Financeiro · Conversa Natural", "0–90 dias"],
        ["FASE 2\nInteligência", C.phase2,  "Operacional · Produção · Compras · RH · Segurança · Chef Digital", "3–9 meses"],
        ["FASE 3\nExpansão", C.phase3,  "Nutrição · Marketing · Delivery · Atendimento · IA Criativa", "9–18 meses"],
        ["FASE 4\nTransformação", C.phase4, "Visão Computacional · Gêmeo Digital · Simulador · IA Cientista", "18–36 meses"],
      ]),

      spacer(100),
      para([run(
        "Critério de priorização: ROI imediato → facilidade de adoção → ausência de barreiras regulatórias → dependências técnicas.",
        { size: 19, color: C.sub, italics: true }
      )]),

      // ── FASE 1 ────────────────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionHeader("FASE 1 — FUNDAÇÃO DE DADOS", "MVP · 0 a 90 dias · Prove valor antes de investir mais", C.green),
      spacer(80),
      para([run(
        "Objetivo: fazer o restaurante sentir o sistema funcionando com dados que já possui. Nenhuma câmera, nenhuma integração complexa. " +
        "Entrada manual ou por planilha no primeiro ciclo. ROI visível em até 30 dias.",
        { size: 21, color: C.text }
      )]),
      spacer(120),

      moduleCard("M5", "Estoque Inteligente", "A base de tudo. Sem dados de estoque, nada mais funciona.", [
        "Cadastro de todos os ingredientes com validade, lote, fornecedor e custo",
        "Cálculo automático de CMV (Custo da Mercadoria Vendida)",
        "Alertas de vencimento e ruptura antes que aconteçam",
        "Registro de desperdício e perdas com rastreabilidade",
        "Previsão de quando comprar e quanto comprar",
      ], "MVP"),
      spacer(120),

      moduleCard("M7", "Engenharia de Cardápio", "ROI imediato e mensurável.", [
        "Análise de CMV, margem e lucro por prato",
        "Matriz Estrela / Vaca / Abacaxi / Enigma (Boston Matrix adaptada)",
        "Cálculo de lucro por minuto de produção e por funcionário",
        "Sugestão de reajuste de preço baseada em dados reais",
        "Identificação de pratos que prejudicam a operação (alto tempo, baixa margem)",
      ], "MVP"),
      spacer(120),

      moduleCard("M11", "Financeiro", "Controle sem planilha manual.", [
        "Fluxo de caixa simples com entradas e saídas",
        "DRE gerencial automático",
        "Ponto de equilíbrio e capital de giro",
        "Previsão financeira dos próximos 30, 60 e 90 dias",
        "Dashboard com EBITDA e margem de contribuição",
      ], "MVP"),
      spacer(120),

      moduleCard("M19", "Conversa Natural", "A interface que dá vida ao sistema.", [
        "O dono pergunta em linguagem comum: 'Como foi hoje?' 'Por que o CMV subiu?'",
        "Respostas objetivas com dados reais do restaurante",
        "Alertas proativos por WhatsApp ou painel",
        "Sem necessidade de abrir relatórios ou planilhas",
        "Base para todos os módulos futuros se comunicarem com o gestor",
      ], "MVP"),

      // ── FASE 2 ────────────────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionHeader("FASE 2 — INTELIGÊNCIA OPERACIONAL", "Diagnóstico · 3 a 9 meses · Conectar dados para prever e agir", C.phase2),
      spacer(80),
      para([run(
        "Objetivo: transformar dados isolados em inteligência conectada. O sistema passa a identificar gargalos, " +
        "prever demanda e dialogar proativamente com o gestor. Aqui nasce o verdadeiro diferencial.",
        { size: 21, color: C.text }
      )]),
      spacer(120),

      moduleCard("M3", "Inteligência Operacional", "Onde está o gargalo? O sistema sabe.", [
        "Mapeamento de tempo de preparo, montagem, embalagem e entrega",
        "Identificação automática de gargalos por estação de trabalho",
        "Respostas para: 'O almoço atrasou por quê?' 'Quem está sobrecarregado?'",
        "Análise de fluxo de pedidos e picos de produção",
        "Comparativo de performance entre turnos e dias da semana",
      ], "Fase2"),
      spacer(120),

      moduleCard("M4", "Previsão de Produção", "Cozinha certa para o dia certo.", [
        "Previsão automática de clientes por histórico, dia, clima e eventos",
        "Quanto arroz, feijão, molho, caldo e insumos produzir por dia",
        "Integração com calendário: feriados, jogos, shows, eventos da cidade",
        "Ajuste automático baseado em padrões aprendidos semana a semana",
        "Redução de sobras e de falta de produto",
      ], "Fase2"),
      spacer(120),

      moduleCard("M6", "Inteligência de Compras", "Comprar certo, do fornecedor certo, na hora certa.", [
        "Análise de todos os fornecedores: preço, prazo, qualidade e histórico",
        "Previsão de quando um produto vai aumentar ou faltar",
        "Alerta de substituição de fornecedor com dados comparativos",
        "Ranking de confiabilidade e custo-benefício por item",
        "Sugestão automática de pedido de compra",
      ], "Fase2"),
      spacer(120),

      moduleCard("M9", "Segurança dos Alimentos", "Conformidade sem esforço manual.", [
        "Monitoramento de temperatura: geladeiras, freezers, fogões, fornos e banho-maria",
        "Alertas de tempo de exposição acima do limite",
        "Checklist digital de APPCC e BPF integrado à rotina",
        "Detecção de risco de contaminação cruzada por processo",
        "Registro automático para auditorias e vigilância sanitária",
      ], "Fase2"),
      spacer(120),

      moduleCard("M10", "RH Inteligente", "Pessoas: o ativo mais sensível do restaurante.", [
        "Avaliação de produtividade por função e turno (sem ranking exposto à equipe)",
        "Identificação de sobrecarga antes do burnout acontecer",
        "Controle de horas extras, assiduidade e pontualidade",
        "Sugestões de treinamento baseadas em gaps de desempenho observados",
        "Nunca demite — apenas sugere ações ao gestor",
      ], "Fase2"),
      spacer(120),

      moduleCard("M20", "Chef Digital — Copiloto do Gestor", "O verdadeiro diferencial do RIOS.", [
        "Não apenas responde: questiona o gestor quando detecta anomalias",
        "'Nas últimas 3 quintas, a produção de arroz foi 18% maior que o consumo. Ajusto a previsão?'",
        "'O cozinheiro Carlos reduziu o tempo de montagem em 11%. Transformo em POP?'",
        "Proativo, contextual e aprende com as decisões do gestor",
        "Transforma o RIOS de ferramenta reativa em parceiro estratégico",
      ], "Fase2"),

      // ── FASE 3 ────────────────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionHeader("FASE 3 — EXPANSÃO DIGITAL", "Alcance · 9 a 18 meses · Marketing, delivery e experiência", C.phase3),
      spacer(80),
      para([run(
        "Objetivo: ampliar a inteligência para além das quatro paredes da cozinha. " +
        "O sistema passa a monitorar a experiência do cliente, o delivery e a presença digital.",
        { size: 21, color: C.text }
      )]),
      spacer(120),

      moduleCard("M8", "Nutrição", "Informação que fideliza e diferencia.", [
        "Cálculo automático de calorias, macros e micronutrientes por prato",
        "Identificação e alerta de alérgenos",
        "Índice glicêmico, sódio, vitaminas e minerais",
        "Geração automática de tabela nutricional para cardápio e delivery",
        "Adequação a dietas especiais (low carb, vegano, sem glúten etc.)",
      ], "Fase3"),
      spacer(120),

      moduleCard("M12", "Marketing Inteligente", "Dados de fora, estratégia de dentro.", [
        "Análise de avaliações: Google, iFood, Instagram e concorrentes",
        "Monitoramento de comentários e sentimento do cliente",
        "Identificação de tendências antes de os concorrentes perceberem",
        "Sugestão de campanhas baseadas em dados reais de vendas e sazonalidade",
        "Comparativo de desempenho digital com restaurantes similares",
      ], "Fase3"),
      spacer(120),

      moduleCard("M13", "Inteligência de Delivery", "Cada minuto perdido é avaliação perdida.", [
        "Análise de tempo de preparo versus tempo de entrega por canal",
        "Identificação de motoboys com melhor desempenho e rotas mais eficientes",
        "Correlação entre embalagem, temperatura na entrega e avaliação",
        "Análise de cancelamentos e reclamações com causa-raiz",
        "Sugestão de janelas de horário para reduzir filas e atrasos",
      ], "Fase3"),
      spacer(120),

      moduleCard("M14", "Inteligência de Atendimento", "O cliente fala — mesmo sem abrir a boca.", [
        "Análise de tempo de espera e chamadas ao garçom por mesa",
        "Identificação de sinais de insatisfação no fluxo do atendimento",
        "Métricas de retorno e fidelização por perfil de cliente",
        "Correlação entre tempo sentado e ticket médio",
        "Sugestões de ajuste de cardápio e atendimento baseadas em comportamento real",
      ], "Fase3"),
      spacer(120),

      moduleCard("M18", "IA Criativa", "Conteúdo e processos gerados automaticamente.", [
        "Criação de receitas novas baseadas nos ingredientes disponíveis em estoque",
        "Geração de combos e promoções com cálculo de margem embutido",
        "Produção automática de POPs, fichas técnicas e treinamentos",
        "Sugestão de fotos, textos e legendas para redes sociais",
        "Criação de cardápios sazonais com base em tendências e histórico de vendas",
      ], "Fase3"),

      // ── FASE 4 ────────────────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionHeader("FASE 4 — VISÃO E TRANSFORMAÇÃO", "Futuro · 18 a 36 meses · Tecnologia de fronteira", C.phase4),
      spacer(80),
      para([run(
        "Objetivo: implementar tecnologias avançadas que requerem escala, dados históricos e aceitação da equipe. " +
        "Estas fases dependem do sucesso das anteriores para serem viáveis — técnica e comercialmente.",
        { size: 21, color: C.text }
      )]),
      spacer(80),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 3, color: C.yellow }, bottom: border0, left: border0, right: border0 },
          shading: { fill: "FFFBEB", type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [para([
            run("Atenção: ", { bold: true, size: 20, color: C.yellow }),
            run("A Fase 4 tem barreiras reais. Visão computacional em cozinhas levanta questões de LGPD e sindicatos. " +
                "O Gêmeo Digital exige integração com planta física e equipamentos. A IA Cientista só existe com dados de centenas de restaurantes. " +
                "Não trate estes módulos como promessas de curto prazo.", { size: 20, color: C.text }),
          ])],
        })]})],
      }),
      spacer(120),

      moduleCard("M1", "Visão Computacional", "A IA que observa a cozinha em tempo real.", [
        "Detecta funcionários, utensílios, ingredientes e equipamentos por câmera",
        "Mapas de calor, fluxo e distância percorrida pela equipe",
        "Identificação de riscos de acidente e uso correto de EPIs",
        "Análise de ergonomia e congestionamentos de circulação",
        "Pré-requisito: LGPD, acordo sindical e treinamento de equipe",
      ], "Fase4"),
      spacer(120),

      moduleCard("M2", "Gêmeo Digital da Cozinha", "Simule antes de mover um único equipamento.", [
        "Modelo virtual completo do restaurante: paredes, bancadas, equipamentos",
        "Simulação de mudanças de layout antes de executar na vida real",
        "Integração com dados do Módulo 1 para validar fluxo simulado",
        "Teste de novos postos de trabalho, equipamentos e configurações",
        "Resultado: '21% menos deslocamento sem comprar nada novo'",
      ], "Fase4"),
      spacer(120),

      moduleCard("M15", "Consultor Estratégico", "O dono pensa em voz alta. A IA responde com dados.", [
        "'Vou abrir outra unidade.' → A IA apresenta riscos, ROI, capital e equipe necessária",
        "Análise de ponto comercial com dados de fluxo e concorrência",
        "Projeções financeiras baseadas no histórico da unidade atual",
        "Avaliação de decisões estratégicas com múltiplos cenários",
        "Pré-requisito: histórico mínimo de 12 meses de operação no RIOS",
      ], "Fase4"),
      spacer(120),

      moduleCard("M16", "Simulador de Cenários", "E se? A pergunta mais valiosa do negócio.", [
        "'Se aumentar 10% o salário?' → Impacto no fluxo de caixa simulado",
        "'Se trocar o fornecedor?' → Comparativo de custo, prazo e risco",
        "'Se contratar mais um cozinheiro?' → Capacidade produtiva e ROI",
        "'Se comprar um forno novo?' → Payback calculado com dados reais",
        "Todas as simulações com intervalos de confiança, não apenas números únicos",
      ], "Fase4"),
      spacer(120),

      moduleCard("M17", "IA Cientista", "O módulo mais revolucionário — e o mais paciente.", [
        "Aprende com dados de TODOS os restaurantes na plataforma",
        "Descobre padrões que nenhum consultor humano conseguiria ver",
        "Exemplo: 'Restaurantes japoneses acima de R$500k desperdiçam 14% menos com esta configuração de ilha fria'",
        "Ninguém programa isso. Ela descobre.",
        "Pré-requisito: centenas de restaurantes ativos na plataforma com dados confiáveis",
      ], "Fase4"),

      // ── PRÓXIMOS PASSOS ───────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionHeader("PRÓXIMOS PASSOS IMEDIATOS", "O que fazer nos próximos 30 dias", C.accent),
      spacer(120),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [700, 2400, 5460, 800],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders: bordersGray, width: { size: 700, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 100 }, children: [para([run("#", { bold: true, size: 19, color: C.white })])] }),
            new TableCell({ borders: bordersGray, width: { size: 2400, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [para([run("AÇÃO", { bold: true, size: 19, color: C.white })])] }),
            new TableCell({ borders: bordersGray, width: { size: 5460, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [para([run("DETALHE", { bold: true, size: 19, color: C.white })])] }),
            new TableCell({ borders: bordersGray, width: { size: 800, type: WidthType.DXA }, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 80, right: 80 }, children: [para([run("PRAZO", { bold: true, size: 19, color: C.white })], { alignment: AlignmentType.CENTER })] }),
          ]}),
          ...[
            ["1", "Definir restaurante piloto", "Escolher 1 restaurante real com faturamento acima de R$50k/mês, dono presente e disposição para testar. Esse primeiro caso vira o case comercial.", "15 dias"],
            ["2", "Mapear dados disponíveis", "Verificar quais dados o piloto já tem: planilhas de estoque, CMV, financeiro, cardápio. Definir o gap de coleta.", "15 dias"],
            ["3", "Definir modelo de negócio", "SaaS mensal? Percentual sobre resultado? Consultoria + tecnologia? Essa decisão muda o roadmap técnico e o argumento de venda.", "30 dias"],
            ["4", "Prototipar M5 + M7 + M19", "MVP funcional: input de dados de estoque e cardápio → sistema calcula CMV e margem → responde perguntas em linguagem natural.", "90 dias"],
            ["5", "Documentar o resultado do piloto", "Capturar: quanto o piloto economizou, o que mudou na operação, qual o ROI real. Isso vira o argumento para os próximos clientes.", "90 dias"],
          ].map(([n, acao, det, prazo]) => new TableRow({ children: [
            new TableCell({ borders: bordersGray, width: { size: 700, type: WidthType.DXA }, shading: { fill: C.accent, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, margins: { top: 100, bottom: 100, left: 120, right: 100 }, children: [para([run(n, { bold: true, size: 20, color: C.white })], { alignment: AlignmentType.CENTER })] }),
            new TableCell({ borders: bordersGray, width: { size: 2400, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [para([run(acao, { bold: true, size: 20, color: C.text })])] }),
            new TableCell({ borders: bordersGray, width: { size: 5460, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [para([run(det, { size: 19, color: C.text })])] }),
            new TableCell({ borders: bordersGray, width: { size: 800, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 100, bottom: 100, left: 80, right: 80 }, children: [para([run(prazo, { size: 18, color: C.sub, bold: true })], { alignment: AlignmentType.CENTER })] }),
          ]})),
        ],
      }),

      spacer(200),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 3, color: C.accent }, bottom: border0, left: border0, right: border0 },
          shading: { fill: C.light, type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 300, right: 300 },
          children: [
            para([run("Uma nota sobre o ritmo", { bold: true, size: 22, color: C.dark })]),
            spacer(60),
            para([run(
              "Este documento é vivo. Cada fase será revisada com base nos resultados reais do piloto. " +
              "O que está escrito aqui é a melhor lógica disponível hoje — mas os dados do restaurante real vão corrigir, " +
              "refinar e talvez surpreender. A vantagem do RIOS é essa: ele aprende. E nós também.",
              { size: 21, color: C.text, italics: true }
            )]),
          ],
        })]})],
      }),

    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/sessions/focused-eloquent-newton/mnt/outputs/RIOS_Roadmap_Estrategico.docx', buffer);
  console.log('OK');
});
);
