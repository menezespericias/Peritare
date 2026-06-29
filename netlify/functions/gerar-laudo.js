// netlify/functions/gerar-laudo.js
// Recebe os dados do processo via POST e devolve o .docx em base64

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, UnderlineType,
  LineRuleType, convertMillimetersToTwip, TableOfContents,
} = require('docx');

// ── cores ──────────────────────────────────────────────────────────────────
const AZUL     = '1F4E79';
const AZUL_MED = '2E75B6';
const AZUL_CLR = 'BDD7EE';
const CINZA    = 'F2F2F2';
const BRANCO   = 'FFFFFF';
const PRETO    = '000000';
const FONTE    = 'Arial';

function ptTwip(pt) { return pt * 20; }

const METODOLOGIA_ITEMS = [
  'a) Minuciosos exames dos documentos questionados (item 5.1);',
  'b) Minuciosos exames dos padrões de confronto (itens 5.2-A e 5.2-B);',
  'c) Cotejos e transcotejos entre documentos questionados e respectivos paradigmas;',
  'd) Utilização de aparelhamento especializado e softwares adequados (Adendo 5);',
  'e) Determinação de convergências e divergências (Item 6);',
  'f) Coordenação dos dados técnicos apurados;',
  'g) Preparação das ilustrações necessárias;',
  'h) Elaboração do laudo pericial grafotécnico;',
  'i) Revisão do laudo;',
  'j) Apresentação do laudo ao juizado competente.',
];

const EOG_NOMES = [
  'Ritmo','Dinamismo','Velocidade','Habilidade','Espontaneidade',
  'Ataques','Remates','Pressão e Progressão','Hábitos Gráficos','Momentos Gráficos',
  'Comportamento de Base','Comportamento de Pauta','Inclinação Axial','Espaçamentos Gráficos','Tendência de Punho',
  'Trajetórias','Calibres','Gladiolagem','Gramas','Proporcionalidade',
  'Sobreposição de Peças','Microscopia','Separação','Cronologia','Vídeos',
  'Imagem do Grafismo','Escavalamento','Cultura Gráfica','Maneirismos Gráficos','Afinamento',
  'Grafogeometria','Traçados das Letras','Polimorfismo Gráfico','Caligrafia','Diversos',
];

// ── helpers ────────────────────────────────────────────────────────────────
function par(children, opts = {}) {
  const arr = typeof children === 'string'
    ? [new TextRun({ text: children, font: FONTE, size: opts.size || 20, bold: opts.bold, color: opts.color || PRETO, italics: opts.italic })]
    : children;
  return new Paragraph({
    children: arr,
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { before: ptTwip(opts.before ?? 4), after: ptTwip(opts.after ?? 4), line: 276, lineRule: LineRuleType.AUTO },
    indent: opts.indent ? { left: convertMillimetersToTwip(opts.indent) } : undefined,
  });
}

function run(text, opts = {}) {
  return new TextRun({ text, font: FONTE, size: opts.size || 20, bold: opts.bold, italics: opts.italic, color: opts.color || PRETO });
}

function h1(text) {
  return new Paragraph({
    text, heading: HeadingLevel.HEADING_1,
    spacing: { before: ptTwip(14), after: ptTwip(6) },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL_MED, space: 2 } },
  });
}
function h2(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: ptTwip(10), after: ptTwip(4) } }); }
function h3(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_3, spacing: { before: ptTwip(8), after: ptTwip(3) } }); }
function pbreak() { return new Paragraph({ children: [new PageBreak()] }); }
function esp(pt = 6) { return new Paragraph({ spacing: { before: ptTwip(pt), after: 0 } }); }

function cel(text, opts = {}) {
  const children = typeof text === 'string'
    ? [par(text, { before: 2, after: 2, bold: opts.bold, color: opts.color, align: opts.align })]
    : text;
  return new TableCell({
    children,
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    columnSpan: opts.span,
    shading: opts.bg ? { fill: opts.bg, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
  });
}

function tblEOG(pqLabel, conv, div, sd, sa, total) {
  const pct = v => ((v / total) * 100).toFixed(2) + '%';
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [
        new TableCell({ children: [par('QUADRO EOG - ELEMENTOS DE ORDEM GERAL', { bold: true, color: BRANCO, align: AlignmentType.CENTER, before: 2, after: 2 })],
          columnSpan: 6, shading: { fill: AZUL, type: ShadingType.CLEAR } }),
      ]}),
      new TableRow({ children: [
        cel('PADRÃO', { bold: true, color: BRANCO, bg: AZUL_MED, width: 12 }),
        cel('CONVERGENTE', { bold: true, color: BRANCO, bg: AZUL_MED, width: 18, align: AlignmentType.CENTER }),
        cel('DIVERGENTE', { bold: true, color: BRANCO, bg: AZUL_MED, width: 18, align: AlignmentType.CENTER }),
        cel('SEM DEFINIR', { bold: true, color: BRANCO, bg: AZUL_MED, width: 18, align: AlignmentType.CENTER }),
        cel('SEM ANÁLISE', { bold: true, color: BRANCO, bg: AZUL_MED, width: 18, align: AlignmentType.CENTER }),
        cel('TOTAL', { bold: true, color: BRANCO, bg: AZUL_MED, width: 16, align: AlignmentType.CENTER }),
      ]}),
      new TableRow({ children: [
        cel(pqLabel, { bold: true, bg: CINZA }),
        cel(String(conv), { align: AlignmentType.CENTER }),
        cel(String(div), { align: AlignmentType.CENTER, bold: true, color: 'C00000' }),
        cel(String(sd), { align: AlignmentType.CENTER }),
        cel(String(sa), { align: AlignmentType.CENTER }),
        cel(String(total), { align: AlignmentType.CENTER, bold: true }),
      ]}),
      new TableRow({ children: [
        cel('EM PERCENTUAIS', { bold: true, bg: CINZA }),
        cel(pct(conv), { align: AlignmentType.CENTER }),
        cel(pct(div), { align: AlignmentType.CENTER, bold: true, color: 'C00000' }),
        cel(pct(sd), { align: AlignmentType.CENTER }),
        cel(pct(sa), { align: AlignmentType.CENTER }),
        cel('100,00%', { align: AlignmentType.CENTER, bold: true }),
      ]}),
    ],
  });
}

// ── gerador principal ──────────────────────────────────────────────────────
async function gerarDocx(dados) {
  const { processo, pqs, paradigmas, metodologia, perito } = dados;

  // cabeçalho
  const header = new Header({
    children: [
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: { top:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE},insideH:{style:BorderStyle.NONE},bottom:{style:BorderStyle.SINGLE,size:6,color:AZUL_MED} },
        rows: [new TableRow({ children: [
          new TableCell({ children: [
            par([run('MENEZES PERÍCIAS', { bold: true, color: AZUL, size: 20 })], { before: 0, after: 0 }),
            par(`Perito: ${perito.nome}`, { before: 0, after: 0, size: 16 }),
            par(perito.registro, { before: 0, after: 0, size: 16 }),
          ], width: { size: 50, type: WidthType.PERCENTAGE }, borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE} } }),
          new TableCell({ children: [
            par(`Processo: ${processo.numero}`, { align: AlignmentType.RIGHT, before: 0, after: 0, size: 16 }),
            par(`Autor: ${processo.autor}`, { align: AlignmentType.RIGHT, before: 0, after: 0, size: 16 }),
            par(`Réu: ${processo.reu}`, { align: AlignmentType.RIGHT, before: 0, after: 0, size: 16 }),
          ], width: { size: 50, type: WidthType.PERCENTAGE }, borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE} } }),
        ]})],
      }),
    ],
  });

  // rodapé
  const footer = new Footer({
    children: [par([
      run(`${perito.nome} – ${perito.registro}    `, { size: 16, color: '666666' }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONTE, size: 16, color: '666666' }),
    ], { align: AlignmentType.CENTER, before: 4, after: 0, border: { top: { style: BorderStyle.SINGLE, size: 4, color: AZUL_MED, space: 2 } } })],
  });

  // ── seções ────────────────────────────────────────────────────────────
  const children = [];

  // CAPA
  children.push(
    par([run(' ', { size: 80 })], { before: 0, after: 0, shading: { fill: AZUL, type: ShadingType.CLEAR } }),
    par([run(' ', { size: 30 })], { before: 0, after: 0, shading: { fill: AZUL_MED, type: ShadingType.CLEAR } }),
    esp(80),
    par([run('MENEZES PERÍCIAS', { bold: true, size: 36, color: AZUL })], { align: AlignmentType.CENTER, before: 0, after: 6 }),
    par([run(perito.registro, { size: 20, color: AZUL_MED })], { align: AlignmentType.CENTER, before: 0, after: 60 }),
    esp(80),
    par([run('LAUDO PERICIAL GRAFOTÉCNICO', { bold: true, size: 48, color: AZUL_MED })], { align: AlignmentType.CENTER, before: 0, after: 20 }),
    esp(120),
    par([run(`Processo nº ${processo.numero}`, { size: 22, color: AZUL })], { align: AlignmentType.CENTER, before: 0, after: 6 }),
    par([run(processo.autor, { size: 20 }), run('  vs  ', { size: 20, color: '666666' }), run(processo.reu, { size: 20 })], { align: AlignmentType.CENTER, before: 0, after: 6 }),
    par([run(`${processo.vara} · Comarca de ${processo.comarca}`, { size: 20, color: '555555' })], { align: AlignmentType.CENTER, before: 0, after: 60 }),
    esp(80),
    par([run(dados.data_laudo || new Date().toLocaleDateString('pt-BR', { day:'numeric', month:'long', year:'numeric' }), { size: 20, color: '555555' })], { align: AlignmentType.CENTER }),
  );

  // SUMÁRIO
  children.push(pbreak(), h1('SUMÁRIO'), new TableOfContents('Sumário', { hyperlink: true, headingStyleRange: '1-3' }));

  // SEÇÃO 1
  children.push(
    pbreak(), h1('1. APRESENTAÇÃO E ENDEREÇAMENTO'),
    par([
      run(`Exmo. Sr. Dr. Juiz de Direito da ${processo.vara} da Comarca de `),
      run(processo.comarca.toUpperCase(), { bold: true }),
    ]),
    esp(6),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [ cel('Processo nº:', { bold: true, bg: CINZA, width: 25 }), cel(processo.numero, { width: 75 }) ] }),
        new TableRow({ children: [ cel('Autor:', { bold: true, bg: CINZA }), cel(processo.autor) ] }),
        new TableRow({ children: [ cel('Réu:', { bold: true, bg: CINZA }), cel(processo.reu) ] }),
        new TableRow({ children: [ cel('Perito Judicial:', { bold: true, bg: CINZA }), cel(perito.nome, { bold: true }) ] }),
        new TableRow({ children: [ cel('Registro:', { bold: true, bg: CINZA }), cel(`${perito.registro}, CPF nº ${perito.cpf}`) ] }),
      ],
    }),
    esp(8),
    par([
      run(perito.nome, { bold: true }),
      run(`, ${perito.titulo}, ${perito.registro}, CPF nº ${perito.cpf}, perito grafotécnico e em documentos, nomeado por este Juízo através do documento id. `),
      run(processo.nomeacao_id || '[ID]', { bold: true }),
      run(`, de `),
      run(processo.nomeacao_data || '[DATA]', { bold: true }),
      run(`, dos autos em referência, vem, respeitosamente, apresentar e submeter à apreciação de Vossa Excelência o incluso Laudo Pericial.`),
    ]),
  );

  // SEÇÃO 2
  children.push(
    h1('2. OBJETIVOS DA PERÍCIA'),
    par('O presente Laudo Pericial tem por finalidade:'),
    par([run('1. Constatar a '), run('autenticidade, ou não,', { bold: true }), run(' dos lançamentos gráficos (assinaturas) questionados, atribuídos ao(à) Autor(a), exarados nos documentos identificados abaixo:')], { indent: 10 }),
    esp(4),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [ cel('Número do Contrato / Documento', { bold: true, color: BRANCO, bg: AZUL }), cel('Fls.', { bold: true, color: BRANCO, bg: AZUL }), cel('Id. Processual', { bold: true, color: BRANCO, bg: AZUL }) ] }),
        ...pqs.map((pq, i) => new TableRow({ children: [
          cel(pq.numero || '-', { bg: i%2===0 ? CINZA : undefined }),
          cel(pq.fls || '-', { align: AlignmentType.CENTER, bg: i%2===0 ? CINZA : undefined }),
          cel(pq.id_proc || '-', { align: AlignmentType.CENTER, bg: i%2===0 ? CINZA : undefined }),
        ]})),
      ],
    }),
    esp(6),
    par([run('2. Responder aos '), run('quesitos formulados pelas partes,', { bold: true }), run(' em conformidade com os exames técnico-científicos realizados.')], { indent: 10 }),
  );

  // SEÇÃO 3
  const metItens = (metodologia?.items || []).map(i => METODOLOGIA_ITEMS[i]).filter(Boolean);
  if (metodologia?.outros) {
    metodologia.outros.split('\n').filter(l=>l.trim()).forEach((l,i) => {
      const letra = String.fromCharCode(107 + metItens.length + i);
      metItens.push(`${letra}) ${l.trim()}`);
    });
  }
  children.push(
    h1('3. METODOLOGIA GERAL'),
    par('A metodologia aplicada ao caso consistiu no desenvolvimento dos seguintes itens:'),
    ...metItens.map(it => par(it, { indent: 10, before: 2, after: 2 })),
  );

  // SEÇÃO 4
  children.push(
    h1('4. INTRODUÇÃO'),
    par('É importante salientar que a assinatura é uma marca pessoal usada para comprovar a intenção na realização de transações que envolvam documentos. Dada a sua importância, qualquer suspeita de fraude deve ser avaliada por peritos com o objetivo de verificar a autenticidade ou identificar o autor do lançamento gráfico.'),
    esp(4),
    par([
      run('O trabalho se baseia nas Leis do Grafismo, especialmente no princípio fundamental de que '),
      run('"A escrita é individual e inconfundível, independente do alfabeto utilizado para a sua produção"', { bold: true, italic: true }),
      run(' (Solange Pellat), demonstrando que o personalismo gráfico existente no punho escritor é cientificamente comprovado.'),
    ]),
  );

  // SEÇÃO 5
  children.push(
    h1('5. DOCUMENTOS SUBMETIDOS A EXAME'),
    h2('5.1. Documentos Questionados (Padrão Questionado - PQ)'),
    par('As assinaturas questionadas estão acostadas nos documentos dos autos, conforme ids e descrições constantes abaixo e são trazidos na Seção Anexos do presente laudo.'),
    par([
      run('Os documentos analisados são cópias inseridas nos autos, pelas partes. Dessa forma, em razão da perícia não ter acesso aos originais, não foi possível analisar a pressão/calibre com 100% de precisão. A impossibilidade de análise precisa de pressão é suprida pela existência de divergências fundamentais em elementos dinâmicos e estruturais — '),
      run('velocidade, espontaneidade, inclinação axial e ataques/remates', { bold: true }),
      run(' — suficientes para a determinação segura da autoria.'),
    ]),
  );
  if (dados.doc5_extra) {
    dados.doc5_extra.split('\n\n').filter(p=>p.trim()).forEach(p => children.push(par(p)));
  }
  pqs.forEach((pq, i) => {
    children.push(par([
      run(`• Documento ${i+1}: `, { bold: true }),
      run(`${pq.tipo || 'Documento'}, nº ${pq.numero || '-'}, datado de ${pq.data_doc || '-'}, fls. ${pq.fls || '-'}. [ID. ${pq.id_proc || '-'}]`),
    ], { before: 3, after: 3 }));
  });

  // 5.2 Paradigmas
  children.push(h2('5.2. Documentos Padrão (Paradigmas de Confronto - PC)'));
  children.push(par('Para os exames, foram utilizados os seguintes padrões gráficos (Paradigmas de Confronto):'));
  const pce = paradigmas.filter(p=>p.tipo==='PCE');
  const pca = paradigmas.filter(p=>p.tipo==='PCA');
  if (pce.length + pca.length > 0) {
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [ cel('Padrão / Documento', { bold: true, color: BRANCO, bg: AZUL }), cel('Data', { bold: true, color: BRANCO, bg: AZUL }), cel('Referência', { bold: true, color: BRANCO, bg: AZUL }) ] }),
        ...paradigmas.map((pg, i) => new TableRow({ children: [
          cel(`[${pg.tipo}] ${pg.descricao || '-'}`, { bg: i%2===0 ? CINZA : undefined }),
          cel(pg.data_assinatura || '-', { align: AlignmentType.CENTER, bg: i%2===0 ? CINZA : undefined }),
          cel(pg.referencia || '-', { bg: i%2===0 ? CINZA : undefined }),
        ]})),
      ],
    }));
  }
  if (pca.length > 0) {
    children.push(
      h3('A. Padrões Colhidos no Ato Pericial (PCA)'),
      par([
        run('O Perito signatário procedeu à colheita dos padrões gráficos da Autora '),
        run(processo.autor, { bold: true }),
        run(`. Os padrões coletados estão documentados nos Adendos deste laudo.`),
      ]),
    );
  }
  if (pce.length > 0) {
    children.push(h3('B. Padrões Encontrados nos Autos (PCE)'));
    children.push(par('Concorreram para a realização da grafoscopia comparativa os espécimes de assinaturas contidos nos documentos acostados aos autos descritos no quadro 5.2.'));
  }

  // SEÇÃO 6 – um bloco por PQ
  children.push(h1('6. EXAMES PERICIAIS E METODOLOGIA'));
  children.push(h2('6.1. Análise dos Paradigmas'));
  children.push(par([
    run('A análise inicial teve como finalidade a apreensão das '),
    run('características de valor grafocinético e grafométrico', { bold: true }),
    run(' das assinaturas paradigmais.'),
  ]));
  children.push(h2('6.2. Confronto Grafoscópico'));
  children.push(par('Após o estudo dos padrões (PC), o Perito passou a confrontá-los em relação aos lançamentos gráficos questionados (PQ), utilizando instrumentos óticos especiais e aplicando as técnicas específicas da Grafoscopia.'));

  pqs.forEach((pq, i) => {
    const eog = pq.eog || [];
    const cnt = { C:0, D:0, SD:0, SA:0 };
    eog.forEach(e => { cnt[e.resultado || e.res || 'SA']++; });
    const conc = cnt.D > cnt.C ? 'DIVERGÊNCIA' : cnt.C > cnt.D ? 'CONVERGÊNCIA' : 'INCONCLUSIVO';
    const concCor = conc === 'DIVERGÊNCIA' ? 'C00000' : conc === 'CONVERGÊNCIA' ? '14532d' : '7c5c00';

    children.push(
      h3(`6.2.${i+1}. PQ${i+1} - ${pq.tipo || 'Documento'}, nº ${pq.numero || '-'}, datado de ${pq.data_doc || '-'}. [ID. ${pq.id_proc || '-'}]`),
    );

    if (pq.consideracoes) {
      children.push(par(pq.consideracoes));
    }

    children.push(
      par(`A seguir o quadro resumo da análise dos pontos relacionados ao PQ${i+1} em relação aos paradigmas existentes (PCE e PCA).`),
      esp(4),
      tblEOG(`PQ${i+1}`, cnt.C, cnt.D, cnt.SD, cnt.SA, 35),
      esp(8),
      par([
        run('Com base na confrontação dos PCE/PCA (paradigmas constantes dos autos e colhidos) e PQ (paradigmas contestados), o resultado apurado é de '),
        run(conc, { bold: true, color: concCor }),
        run('.'),
      ]),
      esp(10),
    );
  });

  // SEÇÃO 7
  const donePQs = pqs.filter(pq => pq.status === 'done');
  const totalConv = donePQs.reduce((a, pq) => a + (pq.eog || []).filter(e => (e.resultado||e.res) === 'C').length, 0);
  const totalDiv  = donePQs.reduce((a, pq) => a + (pq.eog || []).filter(e => (e.resultado||e.res) === 'D').length, 0);
  const conclusaoGeral = totalDiv > totalConv ? 'DIVERGÊNCIA' : totalConv > totalDiv ? 'CONVERGÊNCIA' : 'INCONCLUSIVO';

  children.push(
    h1('7. CONCLUSÃO'),
    par([run('Após confrontar os grafismos padrões (PCE/PCA) com os grafismos questionados (PQ), e com base nos resultados alcançados ao final dos exames, o signatário é levado a '), run('CONCLUIR', { bold: true }), run(' que:')]),
    esp(8),
    new Paragraph({
      children: [run(`As assinaturas são ${concluao_texto(conclusaoGeral)}`, { bold: true, size: 24, color: 'C00000' })],
      alignment: AlignmentType.CENTER,
      spacing: { before: ptTwip(8), after: ptTwip(8) },
      shading: { fill: 'FFF2CC', type: ShadingType.CLEAR },
      border: { top:{style:BorderStyle.SINGLE,size:8,color:'C00000'}, bottom:{style:BorderStyle.SINGLE,size:8,color:'C00000'}, left:{style:BorderStyle.SINGLE,size:8,color:'C00000'}, right:{style:BorderStyle.SINGLE,size:8,color:'C00000'} },
    }),
    esp(8),
    par([
      run('A assinatura de '),
      run(processo.autor, { bold: true }),
      run(', aposta nos documentos listados no item 5.1 deste laudo, '),
      run('NÃO PROMANARAM', { bold: true }),
      run(' do punho escritor da Autora, sendo, portanto, '),
      run('FALSAS', { bold: true, color: 'C00000' }),
      run('.'),
    ]),
    par('Foram constatadas divergências grafocinéticas significativas, indicando a presença de traços incompatíveis com a gênese gráfica da autora, compatíveis com gesto gráfico alheio ao punho escritor avaliado.'),
  );

  // SEÇÃO 8
  children.push(
    h1('8. RESPOSTA AOS QUESITOS'),
    h2('8.1. Quesitos da Parte Autora'),
    par(dados.quesitos_autor || 'Não foram apresentados quesitos pela parte autora.'),
    h2('8.2. Quesitos da Parte Ré'),
    par('[Inserir respostas aos quesitos da parte ré conforme formulados nos autos]'),
  );

  // SEÇÃO 9
  const hoje = dados.data_laudo || new Date().toLocaleDateString('pt-BR', { day:'numeric', month:'long', year:'numeric' });
  children.push(
    h1('9. ENCERRAMENTO'),
    par(`Nada mais havendo a relatar, foi encerrado o presente laudo, elaborado em via única, para os efeitos legais pertinentes.`),
    esp(30),
    par([run(`${processo.comarca || '[COMARCA]'}, ${hoje}.`)], { align: AlignmentType.CENTER }),
    esp(30),
    new Paragraph({ children: [run(perito.nome, { bold: true })], alignment: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 4, color: PRETO, space: 4 } } }),
    par(perito.titulo, { align: AlignmentType.CENTER, before: 2, after: 2 }),
    par(perito.registro, { align: AlignmentType.CENTER, before: 0 }),
  );

  // ADENDOS
  const adendos = paradigmas.filter(pg => pg.tipo === 'PCA' && pg.arquivo_path);
  children.push(
    pbreak(), h1('ADENDOS'),
    par('Os adendos a seguir compõem parte integrante deste laudo pericial.'),
    esp(8),
  );
  adendos.forEach((ad, i) => {
    children.push(h2(`Adendo ${i+1} – ${ad.descricao}`));
    children.push(par(`[Documento: ${ad.arquivo_path || 'arquivo não disponível'}]`, { italic: true, color: '888888' }));
    children.push(esp(8));
  });

  // ANEXOS
  children.push(
    pbreak(), h1('ANEXOS'),
    par('Os documentos a seguir são os originais (ou cópias digitalizadas) dos documentos submetidos a exame pericial.'),
    esp(8),
  );
  pqs.forEach((pq, i) => {
    children.push(h2(`Documento ${i+1}: ${pq.tipo || 'Documento'}, nº ${pq.numero || '-'}, datado de ${pq.data_doc || '-'}, fls. ${pq.fls || '-'}. [ID. ${pq.id_proc || '-'}]`));
    children.push(par(`[Inserir imagem/PDF do documento PQ${i+1} aqui]`, { italic: true, color: '888888' }));
    children.push(esp(10));
  });

  // ── montar documento ────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONTE, size: 20 }, paragraph: { spacing: { line: 276, lineRule: LineRuleType.AUTO } } },
      },
      paragraphStyles: [
        { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal',
          run: { bold:true, size:28, color:AZUL, font:FONTE },
          paragraph: { spacing:{before:ptTwip(14),after:ptTwip(6)}, outlineLevel:0 } },
        { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal',
          run: { bold:true, size:24, color:AZUL_MED, font:FONTE },
          paragraph: { spacing:{before:ptTwip(10),after:ptTwip(4)}, outlineLevel:1 } },
        { id:'Heading3', name:'Heading 3', basedOn:'Normal', next:'Normal',
          run: { bold:true, size:22, color:AZUL, font:FONTE, italics:true },
          paragraph: { spacing:{before:ptTwip(8),after:ptTwip(3)}, outlineLevel:2 } },
      ],
    },
    sections: [{
      properties: { page: { margin: { top: convertMillimetersToTwip(25), bottom: convertMillimetersToTwip(20), left: convertMillimetersToTwip(30), right: convertMillimetersToTwip(20) } } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  return await Packer.toBase64String(doc);
}

function concluao_texto(conc) {
  if (conc === 'DIVERGÊNCIA') return 'FALSAS (Falsidade Gráfica/Falsificação)';
  if (conc === 'CONVERGÊNCIA') return 'AUTÊNTICAS';
  return 'INCONCLUSIVAS';
}

// ── handler Netlify ─────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }
  try {
    const dados = JSON.parse(event.body);
    const base64 = await gerarDocx(dados);
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ docx: base64, filename: `laudo_${dados.processo.numero.replace(/\D/g,'_')}.docx` }),
    };
  } catch (err) {
    console.error(err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
