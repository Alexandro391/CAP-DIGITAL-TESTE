const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Paleta FMC ──
const TEAL   = '2A5A54';
const TEAL_L = '3A7A72';
const TEAL_BG= 'EAF4F2';
const SLATE  = '334155';
const WHITE  = 'FFFFFF';
const GRAY   = 'F1F5F9';
const RED    = 'DC2626';
const AMBER  = 'D97706';
const GREEN  = '16A34A';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cellMargins = { top: 100, bottom: 100, left: 140, right: 140 };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 6 } },
    children: [new TextRun({ text, bold: true, size: 28, color: TEAL, font: 'Arial' })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, color: TEAL_L, font: 'Arial' })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, color: SLATE, font: 'Arial' })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, size: 22, color: SLATE, font: 'Arial', ...opts })]
  });
}

function bold(text) {
  return new TextRun({ text, bold: true, size: 22, color: SLATE, font: 'Arial' });
}

function code(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 360 },
    shading: { fill: GRAY, type: ShadingType.CLEAR },
    children: [new TextRun({ text, font: 'Courier New', size: 20, color: '1E293B' })]
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, color: SLATE, font: 'Arial' })]
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, color: SLATE, font: 'Arial' })]
  });
}

function alertBox(label, text, color) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 9026, type: WidthType.DXA },
            shading: { fill: color === RED ? 'FEF2F2' : color === AMBER ? 'FFFBEB' : TEAL_BG, type: ShadingType.CLEAR },
            margins: cellMargins,
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({ text: label + ' ', bold: true, size: 20, color, font: 'Arial' }),
                  new TextRun({ text, size: 20, color: SLATE, font: 'Arial' })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
}

function spacer(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({ children: [new TextRun('')], spacing: { before: 0, after: 80 } }));
}

function tableHeader(cols, widths) {
  return new TableRow({
    tableHeader: true,
    children: cols.map((col, i) =>
      new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: TEAL, type: ShadingType.CLEAR },
        margins: cellMargins,
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [new TextRun({ text: col, bold: true, size: 20, color: WHITE, font: 'Arial' })]
        })]
      })
    )
  });
}

function tableRow(cells, widths, shade = false) {
  return new TableRow({
    children: cells.map((cell, i) =>
      new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: shade ? GRAY : WHITE, type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({
          children: [new TextRun({ text: cell, size: 20, color: SLATE, font: 'Arial' })]
        })]
      })
    )
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ═══════════════════════════════
// DOCUMENTO
// ═══════════════════════════════
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: TEAL },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: TEAL_L },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, font: 'Arial', color: SLATE },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: { size: 9206, type: WidthType.DXA },
            columnWidths: [5000, 4206],
            rows: [new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  width: { size: 5000, type: WidthType.DXA },
                  margins: { top: 60, bottom: 60, left: 0, right: 0 },
                  children: [new Paragraph({
                    children: [
                      new TextRun({ text: '.fmc', bold: true, size: 22, color: TEAL, font: 'Arial' }),
                      new TextRun({ text: '  |  CAP Digital — Manual de Implementação', size: 20, color: SLATE, font: 'Arial' })
                    ]
                  })]
                }),
                new TableCell({
                  borders: noBorders,
                  width: { size: 4206, type: WidthType.DXA },
                  margins: { top: 60, bottom: 60, left: 0, right: 0 },
                  children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: 'Confidencial · TI / FMC · 2026', size: 18, color: '94A3B8', font: 'Arial' })]
                  })]
                })
              ]
            })]
          }),
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
            children: [new TextRun('')]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1', space: 4 } },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: 'Sistema de Atendimento CAP · Faculdade de Medicina de Campos', size: 18, color: '94A3B8', font: 'Arial' }),
              new TextRun({ text: '\tPágina ', size: 18, color: '94A3B8', font: 'Arial' }),
              new PageNumber({ size: 18, color: '94A3B8', font: 'Arial' })
            ]
          })
        ]
      })
    },
    children: [

      // ════════════════════════════
      // CAPA
      // ════════════════════════════
      ...spacer(3),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: '.fmc', bold: true, size: 52, color: TEAL, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: 'Faculdade de Medicina de Campos', size: 24, color: SLATE, font: 'Arial' })]
      }),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            shading: { fill: TEAL, type: ShadingType.CLEAR },
            margins: { top: 400, bottom: 400, left: 500, right: 500 },
            width: { size: 9026, type: WidthType.DXA },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'MANUAL DE IMPLEMENTAÇÃO', bold: true, size: 36, color: WHITE, font: 'Arial' })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 120 },
                children: [new TextRun({ text: 'CAP Digital — Sistema de Atendimento', size: 26, color: 'B8DEDA', font: 'Arial' })]
              })
            ]
          })]
        })]
      }),
      ...spacer(2),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Central de Apoio Pedagógico — FMC', size: 22, color: SLATE, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80 },
        children: [new TextRun({ text: 'Versão 1.0  ·  Junho de 2026', size: 20, color: '64748B', font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40 },
        children: [new TextRun({ text: 'Elaborado por: Alexandro Macabu de Lima', size: 20, color: '64748B', font: 'Arial' })]
      }),
      ...spacer(2),
      alertBox('ℹ️ Classificação:', 'Este documento é de uso interno restrito à equipe de TI da FMC. Não distribuir externamente.', TEAL_L),

      pageBreak(),

      // ════════════════════════════
      // 1. VISÃO GERAL
      // ════════════════════════════
      h1('1. Visão Geral do Sistema'),
      p('O CAP Digital é o sistema de atendimento e chamados da Central de Apoio Pedagógico da FMC. Permite que professores e funcionários registrem solicitações de suporte técnico e pedagógico em salas de aula, com notificações automáticas para a equipe via Telegram e acompanhamento em tempo real pelo dashboard administrativo.'),

      ...spacer(1),
      h2('1.1 Arquitetura'),
      p('O sistema é composto por três camadas independentes:'),
      ...spacer(1),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 2400, 4426],
        rows: [
          tableHeader(['Camada', 'Tecnologia', 'Função'], [2200, 2400, 4426]),
          tableRow(['Frontend', 'HTML + JavaScript', 'Formulário público (index.html) e dashboard administrativo (dashboard.html)'], [2200, 2400, 4426]),
          tableRow(['Backend', 'Google Apps Script', 'Recebe chamados, grava na planilha e dispara notificações'], [2200, 2400, 4426], true),
          tableRow(['Dados', 'Google Sheets', 'Banco de dados dos chamados com histórico completo'], [2200, 2400, 4426]),
          tableRow(['Notificações', 'Telegram Bot API', 'Alerta automático no grupo da equipe CAP a cada novo chamado'], [2200, 2400, 4426], true),
          tableRow(['Hospedagem', 'GitHub Pages', 'Servindo os arquivos HTML sem custo e com deploy automático via Git'], [2200, 2400, 4426]),
        ]
      }),

      ...spacer(1),
      h2('1.2 Fluxo de um Chamado'),
      ...spacer(1),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            shading: { fill: TEAL_BG, type: ShadingType.CLEAR },
            margins: cellMargins,
            width: { size: 9026, type: WidthType.DXA },
            children: [
              new Paragraph({
                children: [new TextRun({ text: 'Usuário preenche o formulário  →  Apps Script recebe  →  Grava no Sheets  →  Notifica Telegram  →  Equipe CAP visualiza no Dashboard  →  Atualiza status  →  WhatsApp notifica o solicitante (opcional)', size: 20, color: TEAL, font: 'Arial', italics: true })]
              })
            ]
          })]
        })]
      }),

      ...spacer(1),
      h2('1.3 Repositório e URL'),
      bullet('Repositório GitHub: github.com/Alexandro391/Sistema-de-Atendimento-CAP'),
      bullet('URL de acesso (formulário): alexandro391.github.io/Sistema-de-Atendimento-CAP/'),
      bullet('URL do dashboard: alexandro391.github.io/Sistema-de-Atendimento-CAP/dashboard.html'),
      bullet('Branch principal: main (GitHub Pages configurado nesta branch)'),

      pageBreak(),

      // ════════════════════════════
      // 2. PRÉ-REQUISITOS
      // ════════════════════════════
      h1('2. Pré-Requisitos'),

      h2('2.1 Contas e Acessos Necessários'),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 2200, 3826],
        rows: [
          tableHeader(['Recurso', 'Tipo de Conta', 'Finalidade'], [3000, 2200, 3826]),
          tableRow(['GitHub', 'Gratuita', 'Hospedagem dos arquivos HTML via GitHub Pages'], [3000, 2200, 3826]),
          tableRow(['Google (pessoal ou institucional)', 'Qualquer', 'Google Sheets + Apps Script'], [3000, 2200, 3826], true),
          tableRow(['Telegram', 'Conta pessoal', 'Criação e administração do bot de notificações'], [3000, 2200, 3826]),
        ]
      }),

      ...spacer(1),
      h2('2.2 Ferramentas para a Equipe de TI'),
      bullet('Navegador moderno (Chrome, Firefox ou Edge)'),
      bullet('Acesso ao repositório GitHub (leitura ou escrita, conforme papel)'),
      bullet('Acesso à planilha Google Sheets dos chamados'),
      bullet('Telegram instalado (para gerenciar o bot e o grupo)'),
      ...spacer(1),
      alertBox('⚠️ Atenção:', 'O sistema não requer servidor, banco de dados relacional, Node.js ou qualquer infraestrutura adicional. Toda a lógica roda no navegador e no Google Apps Script.', AMBER),

      pageBreak(),

      // ════════════════════════════
      // 3. ESTRUTURA DE ARQUIVOS
      // ════════════════════════════
      h1('3. Estrutura de Arquivos'),
      p('O repositório contém os seguintes arquivos principais:'),
      ...spacer(1),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          tableHeader(['Arquivo', 'Descrição'], [2800, 6226]),
          tableRow(['index.html', 'Formulário público de abertura de chamados (PWA instalável)'], [2800, 6226]),
          tableRow(['dashboard.html', 'Painel administrativo com gráficos, KPIs, tabela de chamados e controle de status'], [2800, 6226], true),
          tableRow(['conferenciacap.html', 'Tela de conferência/consulta de chamados'], [2800, 6226]),
          tableRow(['manifest.json', 'Configuração PWA para instalação no celular como app'], [2800, 6226], true),
          tableRow(['sw.js', 'Service Worker — permite funcionamento offline e cache do PWA'], [2800, 6226]),
          tableRow(['icon-192.png / icon-512.svg', 'Ícones do app em diferentes resoluções'], [2800, 6226], true),
          tableRow(['favicon.ico', 'Ícone da aba do navegador'], [2800, 6226]),
        ]
      }),

      ...spacer(1),
      alertBox('ℹ️ Nota:', 'Todos os arquivos devem estar na raiz do repositório para que o GitHub Pages sirva corretamente.', TEAL_L),

      pageBreak(),

      // ════════════════════════════
      // 4. CONFIGURAÇÃO DO GOOGLE SHEETS
      // ════════════════════════════
      h1('4. Configuração do Google Sheets'),

      h2('4.1 Criar a Planilha'),
      numbered('Acesse sheets.google.com com a conta desejada'),
      numbered('Crie uma nova planilha e nomeie como CAP Digital - Chamados'),
      numbered('Na Aba1 (renomeie para Página1), insira os cabeçalhos na linha 1 exatamente como abaixo:'),
      ...spacer(1),
      code('A1: timestamp   B1: protocolo   C1: sala   D1: categoria'),
      code('E1: urgencia    F1: descricao   G1: solicitante   H1: tipo'),
      code('I1: telefone    J1: status'),
      ...spacer(1),
      numbered('Copie o ID da planilha da URL do navegador:'),
      code('https://docs.google.com/spreadsheets/d/  [ESTE_TRECHO_É_O_ID]  /edit'),

      ...spacer(1),
      h2('4.2 Tornar a Planilha Pública para Leitura'),
      p('O dashboard lê os dados diretamente da planilha via API pública do Google. Para isso:'),
      numbered('Na planilha, clique em Compartilhar (canto superior direito)'),
      numbered('Em "Acesso geral", selecione Qualquer pessoa com o link'),
      numbered('Defina o papel como Leitor'),
      numbered('Clique em Concluído'),
      ...spacer(1),
      alertBox('⚠️ Atenção:', 'Essa configuração deixa os dados dos chamados visíveis publicamente para quem tiver o link da planilha. Não inclua dados sensíveis além dos previstos no formulário.', AMBER),

      pageBreak(),

      // ════════════════════════════
      // 5. APPS SCRIPT
      // ════════════════════════════
      h1('5. Configuração do Google Apps Script'),

      h2('5.1 Acessar o Editor'),
      numbered('Com a planilha aberta, clique em Extensões → Apps Script'),
      numbered('Apague todo o código existente'),
      numbered('Cole o código fornecido pela equipe de desenvolvimento'),
      numbered('Salve com Ctrl+S ou pelo ícone de disquete'),

      ...spacer(1),
      h2('5.2 Publicar como Aplicativo Web'),
      numbered('Clique em Implantar → Nova implantação'),
      numbered('Em "Tipo", selecione Aplicativo da Web'),
      numbered('Configure:'),
      bullet('Executar como: Eu (conta Google do administrador)', 1),
      bullet('Quem pode acessar: Qualquer pessoa', 1),
      numbered('Clique em Implantar'),
      numbered('Autorize as permissões solicitadas (acesso à planilha)'),
      numbered('Copie a URL gerada — ela terá o formato:'),
      code('https://script.google.com/macros/s/XXXXXXXXXXXXXXXX/exec'),
      ...spacer(1),
      alertBox('⚠️ Atenção:', 'Sempre que o código do Apps Script for alterado, é necessário publicar uma nova implantação (não editar a existente) para que as mudanças entrem em vigor.', AMBER),

      ...spacer(1),
      h2('5.3 Testar o Apps Script'),
      p('Para verificar se o script está funcionando, acesse a URL publicada no navegador. A resposta deve ser um JSON como:'),
      code('{"status":"ok","message":"CAP Apps Script ativo"}'),

      pageBreak(),

      // ════════════════════════════
      // 6. TELEGRAM BOT
      // ════════════════════════════
      h1('6. Configuração do Bot Telegram'),

      h2('6.1 Criar o Bot (se necessário)'),
      p('O sistema já possui um bot configurado (CapChamadosBot). Este processo só é necessário se for preciso recriar o bot.'),
      numbered('Abra o Telegram e busque por @BotFather'),
      numbered('Envie o comando /newbot'),
      numbered('Defina o nome do bot: CAP Chamados FMC'),
      numbered('Defina o username: CapChamadosBot (ou outro disponível)'),
      numbered('Copie o token fornecido — formato: 1234567890:ABCdefGHIjklMNOpqrsTUVwxyz'),

      ...spacer(1),
      h2('6.2 Obter o Chat ID do Grupo'),
      numbered('Adicione o bot ao grupo da equipe CAP'),
      numbered('Envie qualquer mensagem no grupo'),
      numbered('Acesse no navegador:'),
      code('https://api.telegram.org/bot[TOKEN]/getUpdates'),
      numbered('No JSON retornado, localize o campo "chat" → "id" (número negativo para grupos)'),
      numbered('Anote este chat_id'),

      ...spacer(1),
      h2('6.3 Verificar Integração'),
      p('O Apps Script contém a lógica de envio das notificações. Basta verificar que o token e o chat_id no script estejam corretos. Ao abrir um chamado de teste, a notificação deve chegar no grupo do Telegram em segundos.'),

      pageBreak(),

      // ════════════════════════════
      // 7. CONFIGURAÇÃO DOS ARQUIVOS HTML
      // ════════════════════════════
      h1('7. Configuração dos Arquivos HTML'),

      h2('7.1 index.html — Formulário Público'),
      p('Localize e substitua os seguintes valores no arquivo:'),
      ...spacer(1),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2400, 3200, 3426],
        rows: [
          tableHeader(['Variável', 'Localização', 'O que inserir'], [2400, 3200, 3426]),
          tableRow(['APPS_SCRIPT_URL', 'Linha ~1996', 'URL publicada do Apps Script (passo 5.2)'], [2400, 3200, 3426]),
        ]
      }),

      ...spacer(1),
      h2('7.2 dashboard.html — Painel Administrativo'),
      ...spacer(1),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2400, 3200, 3426],
        rows: [
          tableHeader(['Variável', 'Localização', 'O que inserir'], [2400, 3200, 3426]),
          tableRow(['SHEET_ID', 'Linha ~858', 'ID da planilha Google Sheets (passo 4.1)'], [2400, 3200, 3426]),
          tableRow(['APPS_URL', 'Linha ~999', 'URL publicada do Apps Script (passo 5.2)'], [2400, 3200, 3426], true),
          tableRow(['PIN_CODE', 'Linha ~860', 'Código PIN de acesso ao dashboard (padrão: 1234 — alterar!)'], [2400, 3200, 3426]),
        ]
      }),

      ...spacer(1),
      alertBox('🔒 Segurança:', 'O PIN padrão é 1234. Altere imediatamente para um código seguro antes de disponibilizar o dashboard para a equipe. O PIN fica visível no código-fonte — para maior segurança futura, considere autenticação via Firebase Authentication.', RED),

      pageBreak(),

      // ════════════════════════════
      // 8. DEPLOY NO GITHUB PAGES
      // ════════════════════════════
      h1('8. Deploy no GitHub Pages'),

      h2('8.1 Configurar o Repositório'),
      numbered('Acesse github.com e faça login na conta institucional (ou pessoal do administrador)'),
      numbered('Abra o repositório Sistema-de-Atendimento-CAP'),
      numbered('Vá em Settings → Pages'),
      numbered('Em "Source", selecione Deploy from a branch'),
      numbered('Selecione a branch main e a pasta / (root)'),
      numbered('Clique em Save'),
      numbered('Aguarde cerca de 1-2 minutos e acesse a URL gerada'),

      ...spacer(1),
      h2('8.2 Atualizar Arquivos'),
      p('Para atualizar qualquer arquivo do sistema:'),
      numbered('Acesse o repositório no GitHub'),
      numbered('Clique no arquivo desejado (ex: dashboard.html)'),
      numbered('Clique no ícone de lápis (editar)'),
      numbered('Faça as alterações necessárias'),
      numbered('Role até o final e clique em Commit changes'),
      numbered('O GitHub Pages atualiza automaticamente em 1-3 minutos'),
      ...spacer(1),
      alertBox('ℹ️ Dica:', 'Para alterações mais frequentes ou trabalho em equipe, recomenda-se clonar o repositório localmente com Git e usar o fluxo git pull → editar → git push.', TEAL_L),

      pageBreak(),

      // ════════════════════════════
      // 9. MIGRAÇÃO PARA SERVIDOR PRÓPRIO
      // ════════════════════════════
      h1('9. Migração para Servidor da FMC (Opcional)'),

      p('O sistema HTML pode ser migrado para o servidor institucional da FMC sem qualquer alteração no código. A lógica de dados (Apps Script, Sheets, Telegram) permanece inalterada.'),

      ...spacer(1),
      h2('9.1 Requisitos do Servidor'),
      bullet('Apache ou Nginx (qualquer versão recente)'),
      bullet('Acesso externo à internet (para comunicação com Google APIs e Telegram)'),
      bullet('Suporte a HTTPS (obrigatório para PWA e APIs externas)'),
      bullet('Capacidade de servir arquivos estáticos HTML/CSS/JS'),

      ...spacer(1),
      h2('9.2 Procedimento de Migração'),
      numbered('Solicitar ao TI o acesso FTP/SFTP ou painel administrativo (cPanel, Plesk, etc.)'),
      numbered('Copiar todos os arquivos do repositório para a pasta pública do servidor (geralmente public_html/ ou www/)'),
      numbered('Configurar subdomínio: ex. cap.fmc.edu.br apontando para a pasta'),
      numbered('Testar o acesso pela nova URL'),
      numbered('Atualizar o endereço nos materiais de divulgação internos'),
      ...spacer(1),
      alertBox('ℹ️ Importante:', 'Não é necessário instalar Node.js, PHP ou qualquer runtime no servidor. Os arquivos são HTML puro.', TEAL_L),

      pageBreak(),

      // ════════════════════════════
      // 10. PWA — INSTALAÇÃO NO CELULAR
      // ════════════════════════════
      h1('10. PWA — Instalação como Aplicativo'),

      p('O formulário de chamados (index.html) é um Progressive Web App (PWA) e pode ser instalado no celular como um aplicativo nativo, sem loja de aplicativos.'),

      ...spacer(1),
      h2('10.1 Instalar no Android'),
      numbered('Abra o Chrome no celular'),
      numbered('Acesse a URL do sistema'),
      numbered('Toque nos três pontos (menu) → Adicionar à tela inicial'),
      numbered('Confirme o nome e toque em Adicionar'),
      numbered('O ícone da FMC aparecerá na tela inicial'),

      ...spacer(1),
      h2('10.2 Instalar no iOS (iPhone)'),
      numbered('Abra o Safari no iPhone'),
      numbered('Acesse a URL do sistema'),
      numbered('Toque no ícone de compartilhar (quadrado com seta)'),
      numbered('Role e toque em Adicionar à Tela de Início'),
      numbered('Confirme o nome e toque em Adicionar'),

      ...spacer(1),
      h2('10.3 Arquivos PWA Necessários'),
      p('Para que o PWA funcione corretamente, os seguintes arquivos devem estar na raiz do repositório:'),
      bullet('manifest.json — configuração do app (nome, ícone, cor)'),
      bullet('sw.js — service worker para cache e funcionamento offline'),
      bullet('icon-192.png — ícone 192x192 pixels'),
      bullet('icon-512.png — ícone 512x512 pixels'),

      pageBreak(),

      // ════════════════════════════
      // 11. MANUTENÇÃO
      // ════════════════════════════
      h1('11. Manutenção e Operação'),

      h2('11.1 Rotinas Recomendadas'),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 2200, 4826],
        rows: [
          tableHeader(['Frequência', 'Tarefa', 'Responsável'], [2000, 2200, 4826]),
          tableRow(['Diária', 'Verificar novos chamados no dashboard', 'Equipe CAP'], [2000, 2200, 4826]),
          tableRow(['Semanal', 'Exportar CSV de chamados para backup', 'Equipe CAP'], [2000, 2200, 4826], true),
          tableRow(['Mensal', 'Verificar espaço utilizado na planilha Google Sheets', 'TI'], [2000, 2200, 4826]),
          tableRow(['Mensal', 'Revisar se o Apps Script está ativo e respondendo', 'TI'], [2000, 2200, 4826], true),
          tableRow(['Semestral', 'Revisar PIN de acesso ao dashboard', 'TI'], [2000, 2200, 4826]),
          tableRow(['Conforme demanda', 'Atualizar salas/categorias no formulário', 'Desenvolvimento'], [2000, 2200, 4826], true),
        ]
      }),

      ...spacer(1),
      h2('11.2 Problemas Comuns e Soluções'),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3200, 5826],
        rows: [
          tableHeader(['Problema', 'Solução'], [3200, 5826]),
          tableRow(['Dashboard não carrega dados', 'Verificar se a planilha está pública (passo 4.2). Acessar sempre pela URL do GitHub Pages, nunca abrindo o arquivo local.'], [3200, 5826]),
          tableRow(['Notificação Telegram parou', 'Acessar o Apps Script e verificar se o token do bot e o chat_id estão corretos. Reautorizar se necessário.'], [3200, 5826], true),
          tableRow(['Formulário não envia chamado', 'Verificar se a variável APPS_SCRIPT_URL no index.html aponta para a URL correta do script publicado.'], [3200, 5826]),
          tableRow(['Apps Script retorna erro de permissão', 'Revogar e reautorizar as permissões em Implantar → Gerenciar implantações.'], [3200, 5826], true),
          tableRow(['GitHub Pages não atualiza', 'Aguardar até 5 minutos. Verificar em Settings → Pages se há erro de build.'], [3200, 5826]),
          tableRow(['PIN do dashboard esquecido', 'Editar o arquivo dashboard.html no GitHub e localizar a variável PIN_CODE para redefinir.'], [3200, 5826], true),
        ]
      }),

      pageBreak(),

      // ════════════════════════════
      // 12. SEGURANÇA
      // ════════════════════════════
      h1('12. Considerações de Segurança'),

      alertBox('🔒 Vulnerabilidade Conhecida:', 'O PIN de acesso ao dashboard está armazenado em texto puro no código-fonte (client-side). Qualquer pessoa com acesso ao repositório ou que inspecione o código pode visualizá-lo. Para uso atual como controle interno da equipe CAP, o nível de risco é aceitável. Para acesso mais amplo, recomenda-se autenticação via Firebase Authentication.', RED),

      ...spacer(1),
      h2('12.1 Boas Práticas Implementadas'),
      bullet('Dados armazenados no Google Sheets com controle de compartilhamento'),
      bullet('Dashboard protegido por PIN numérico'),
      bullet('Apps Script executado sob conta autenticada do Google'),
      bullet('Comunicação via HTTPS em todas as camadas'),
      bullet('Sem armazenamento de senhas de usuários'),

      ...spacer(1),
      h2('12.2 Melhorias de Segurança Recomendadas (Futuras)'),
      bullet('Substituir PIN por autenticação Firebase (email/senha institucional)'),
      bullet('Restringir leitura da planilha apenas ao domínio @fmc.edu.br'),
      bullet('Implementar rate limiting no Apps Script para evitar spam'),
      bullet('Migrar status dos chamados do localStorage para o Firestore para sincronização entre dispositivos'),

      pageBreak(),

      // ════════════════════════════
      // 13. CONTATOS E SUPORTE
      // ════════════════════════════
      h1('13. Contatos e Suporte'),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2600, 2200, 4226],
        rows: [
          tableHeader(['Responsável', 'Área', 'Contato'], [2600, 2200, 4226]),
          tableRow(['Alexandro Macabu de Lima', 'Desenvolvimento / CAP', 'xand391@gmail.com'], [2600, 2200, 4226]),
          tableRow(['Equipe TI FMC', 'Infraestrutura', 'A definir pela TI'], [2600, 2200, 4226], true),
        ]
      }),

      ...spacer(1),
      h2('13.1 Repositório e Documentação'),
      bullet('Repositório: github.com/Alexandro391/Sistema-de-Atendimento-CAP'),
      bullet('Issues e sugestões: github.com/Alexandro391/Sistema-de-Atendimento-CAP/issues'),

      ...spacer(2),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1', space: 8 } },
        spacing: { before: 200 },
        children: [new TextRun({ text: 'CAP Digital · Faculdade de Medicina de Campos · 2026', size: 18, color: '94A3B8', font: 'Arial' })]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Manual_Implementacao_CAP_Digital.docx', buffer);
  console.log('Manual gerado com sucesso!');
}).catch(err => {
  console.error('Erro:', err);
  process.exit(1);
});
