// app.js
const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static('public'));

// ---------- JSONBin ----------
// Defina essas variáveis no painel da Vercel:
//   Settings → Environment Variables
//   JSONBIN_ID  = ID do seu bin  (ex: 6849f2...)
//   JSONBIN_KEY = Master Key     (ex: $2a$10$...)
//
// O bin deve ser criado com o conteúdo inicial: { "data": {} }

const JSONBIN_URL = `https://api.jsonbin.io/v3/b/${process.env.JSONBIN_ID}`;
const JSONBIN_KEY = process.env.JSONBIN_KEY;

async function carregarDados() {
  const res = await fetch(`${JSONBIN_URL}/latest`, {
    headers: { 'X-Master-Key': JSONBIN_KEY }
  });
  if (!res.ok) throw new Error(`JSONBin GET falhou: ${res.status}`);
  const json = await res.json();
  // suporta tanto { data: {...} } quanto o objeto direto
  return json.record?.data ?? json.record ?? {};
}

async function salvarDados(d) {
  const res = await fetch(JSONBIN_URL, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'X-Master-Key': JSONBIN_KEY
    },
    body: JSON.stringify({ data: d })
  });
  if (!res.ok) throw new Error(`JSONBin PUT falhou: ${res.status}`);
}

// ---------- Constantes ----------
const TOTAL_FILEIRAS    = 10;
const TOTAL_BLOCOS      = 3;
const PLANTAS_POR_BLOCO = 10;
const TOTAL = TOTAL_FILEIRAS * TOTAL_BLOCOS * PLANTAS_POR_BLOCO; // 300

function chave(f, b, n) { return `${f}-${b}-${n}`; }
function rotulo(f, b)   { return `F${f}B${b}`; }

// ---------- Cálculos ----------
function calcSpadPlanta(a, m, b) {
  if (a == null || m == null || b == null) return null;
  return Number(((Number(a) + Number(m) + Number(b)) / 3).toFixed(4));
}

function calcIE(de, dpoe, dpe) {
  if (de == null || dpoe == null || dpe == null) return null;
  de = Number(de); dpoe = Number(dpoe); dpe = Number(dpe);
  return Number((((dpe - de) + ((dpoe - de) + (dpoe - dpe)) / 2) / 2).toFixed(4));
}

function areaFolha(l) {
  l = Number(l);
  return 0.708 * l * l - 10.44 * l + 83.4;
}

function calcAreaFoliar(larguras) {
  if (!larguras || larguras.length === 0) return null;
  const soma = larguras.reduce((acc, l) => acc + areaFolha(l), 0);
  return Number((soma / larguras.length).toFixed(4));
}

// ---------- Rotas ----------

app.get('/dados', async (req, res) => {
  try {
    const dados = await carregarDados();
    res.json({ dados, total: TOTAL, registrados: Object.keys(dados).length });
  } catch (e) {
    console.error('/dados:', e.message);
    res.status(500).json({ erro: e.message });
  }
});

app.post('/salvar', async (req, res) => {
  try {
    const {
      fileira, bloco, numBloco,
      altura, folhas,
      spadProtocolo, spadApice, spadMedio, spadBase,
      ieDE, ieDPOE, ieDPE,
      numCachos,
      largurasFolhas,
      ausente
    } = req.body;

    const dados = await carregarDados();

    const spadPlanta = ausente ? null : calcSpadPlanta(spadApice, spadMedio, spadBase);
    const ie         = ausente ? null : calcIE(ieDE, ieDPOE, ieDPE);
    const areaFoliar = ausente ? null : calcAreaFoliar(largurasFolhas);

    dados[chave(fileira, bloco, numBloco)] = {
      fileira:        Number(fileira),
      bloco:          Number(bloco),
      numBloco:       Number(numBloco),
      ausente:        !!ausente,
      altura:         ausente ? null : Number(altura),
      folhas:         ausente ? null : Number(folhas),
      spadProtocolo:  ausente ? null : Number(spadProtocolo),
      spadApice:      ausente ? null : Number(spadApice),
      spadMedio:      ausente ? null : Number(spadMedio),
      spadBase:       ausente ? null : Number(spadBase),
      spadPlanta,
      ieDE:           ausente ? null : Number(ieDE),
      ieDPOE:         ausente ? null : Number(ieDPOE),
      ieDPE:          ausente ? null : Number(ieDPE),
      ie,
      numCachos:      ausente ? null : Number(numCachos),
      largurasFolhas: ausente ? null : (largurasFolhas || []).map(Number),
      areaFoliar,
    };

    await salvarDados(dados);
    res.json({ ok: true, registrados: Object.keys(dados).length, spadPlanta, ie, areaFoliar });
  } catch (e) {
    console.error('/salvar:', e.message);
    res.status(500).json({ erro: e.message });
  }
});

app.delete('/apagar', async (req, res) => {
  try {
    const { fileira, bloco, numBloco } = req.body;
    const dados = await carregarDados();
    delete dados[chave(fileira, bloco, numBloco)];
    await salvarDados(dados);
    res.json({ ok: true });
  } catch (e) {
    console.error('/apagar:', e.message);
    res.status(500).json({ erro: e.message });
  }
});

app.post('/limpar', async (req, res) => {
  try {
    await salvarDados({});
    res.json({ ok: true });
  } catch (e) {
    console.error('/limpar:', e.message);
    res.status(500).json({ erro: e.message });
  }
});

// ---------- Geração do Excel ----------
app.get('/gerar-excel', async (req, res) => {
  try {
    const dados = await carregarDados();
    const workbook = new ExcelJS.Workbook();

    const borda = {
      top:    { style: 'thin' }, bottom: { style: 'thin' },
      left:   { style: 'thin' }, right:  { style: 'thin' }
    };

    const hStyle = (bgArgb) => ({
      font:      { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 },
      fill:      { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
      border:    borda
    });

    const cellStyle = {
      alignment: { horizontal: 'center', vertical: 'middle' }, border: borda
    };
    const calcStyle = {
      alignment: { horizontal: 'center', vertical: 'middle' }, border: borda,
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } },
      font: { italic: true }
    };
    const ausenteStyle = {
      alignment: { horizontal: 'center', vertical: 'middle' }, border: borda,
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } },
      font: { italic: true, color: { argb: 'FF856404' } }
    };

    // ── ABA 1 ──────────────────────────────────
    const sheet1 = workbook.addWorksheet('Avaliação Morfofisiológica');

    const headers1 = [
      'Fileira e Bloco', 'Nº da Planta', 'Altura (m)', 'Nº de Folhas',
      'SPAD Protocolo', 'SPAD Ápice', 'SPAD Médio', 'SPAD Base', 'SPAD Planta*',
      'IE - DE (mm)', 'IE - DPOE (mm)', 'IE - DPE (mm)', 'IE*',
      'Nº de Cachos', 'Área Foliar (cm²)*', 'Observação'
    ];
    const colW1 = [16,13,12,13,16,13,13,13,14,14,16,14,12,13,18,18];
    const headerColors = [
      'FF2D6A4F','FF2D6A4F',
      'FF40916C','FF40916C',
      'FF1B7A4A','FF1B7A4A','FF1B7A4A','FF1B7A4A','FF52B788',
      'FF2D6A4F','FF2D6A4F','FF2D6A4F','FF52B788',
      'FF40916C',
      'FF52B788',
      'FF2D6A4F'
    ];

    headers1.forEach((h, i) => {
      const cell = sheet1.getCell(1, i + 1);
      cell.value = h;
      cell.style = hStyle(headerColors[i]);
      sheet1.getColumn(i + 1).width = colW1[i];
    });
    sheet1.getRow(1).height = 38;

    let row1 = 2;
    for (let f = 1; f <= TOTAL_FILEIRAS; f++) {
      for (let b = 1; b <= TOTAL_BLOCOS; b++) {
        for (let n = 1; n <= PLANTAS_POR_BLOCO; n++) {
          const r = dados[chave(f, b, n)];
          sheet1.getCell(row1, 1).value = rotulo(f, b);
          sheet1.getCell(row1, 2).value = n;

          if (!r) {
            for (let c = 3; c <= 15; c++) sheet1.getCell(row1, c).value = '-';
            sheet1.getCell(row1, 16).value = 'Não registrado';
            for (let c = 1; c <= 16; c++) sheet1.getCell(row1, c).style = cellStyle;
          } else if (r.ausente) {
            for (let c = 3; c <= 15; c++) sheet1.getCell(row1, c).value = '-';
            sheet1.getCell(row1, 16).value = 'Ausente';
            for (let c = 1; c <= 16; c++) sheet1.getCell(row1, c).style = ausenteStyle;
          } else {
            [r.altura, r.folhas, r.spadProtocolo, r.spadApice, r.spadMedio, r.spadBase]
              .forEach((v, i) => {
                sheet1.getCell(row1, 3 + i).value = v ?? '-';
                sheet1.getCell(row1, 3 + i).style = cellStyle;
              });
            sheet1.getCell(row1, 9).value  = r.spadPlanta ?? '-';
            sheet1.getCell(row1, 9).style  = calcStyle;
            [r.ieDE, r.ieDPOE, r.ieDPE].forEach((v, i) => {
              sheet1.getCell(row1, 10 + i).value = v ?? '-';
              sheet1.getCell(row1, 10 + i).style = cellStyle;
            });
            sheet1.getCell(row1, 13).value = r.ie ?? '-';
            sheet1.getCell(row1, 13).style = calcStyle;
            sheet1.getCell(row1, 14).value = r.numCachos ?? '-';
            sheet1.getCell(row1, 14).style = cellStyle;
            sheet1.getCell(row1, 15).value = r.areaFoliar ?? '-';
            sheet1.getCell(row1, 15).style = calcStyle;
            sheet1.getCell(row1, 16).value = '';
            sheet1.getCell(row1, 16).style = cellStyle;
            sheet1.getCell(row1, 1).style  = cellStyle;
            sheet1.getCell(row1, 2).style  = cellStyle;
          }
          row1++;
        }
      }
    }

    row1++;
    const nota1 = sheet1.getCell(row1, 1);
    nota1.value = '* Calculado automaticamente — SPAD Planta: (Ápice+Médio+Base)/3 | IE: {(DPE-DE)+[(DPOE-DE)+(DPOE-DPE)]/2}/2 | Área Foliar: média de [0,708·L²−10,44·L+83,4] para cada folha';
    nota1.font  = { italic: true, size: 9, color: { argb: 'FF555555' } };
    sheet1.mergeCells(row1, 1, row1, 16);

    // ── ABA 2 ──────────────────────────────────
    const sheet2 = workbook.addWorksheet('Resumo por Bloco');

    const headers2 = [
      'Fileira e Bloco',
      'Média Altura (m)', 'Média Nº Folhas',
      'Média SPAD Protocolo', 'Média SPAD Ápice', 'Média SPAD Médio',
      'Média SPAD Base', 'Média SPAD Planta',
      'Média IE', 'Média Nº Cachos', 'Média Área Foliar (cm²)'
    ];
    const colW2 = [16,18,16,22,18,18,18,18,14,18,22];

    headers2.forEach((h, i) => {
      const cell = sheet2.getCell(1, i + 1);
      cell.value = h;
      cell.style = hStyle('FF2D6A4F');
      sheet2.getColumn(i + 1).width = colW2[i];
    });
    sheet2.getRow(1).height = 38;

    const camposMedidas = [
      'altura','folhas',
      'spadProtocolo','spadApice','spadMedio','spadBase','spadPlanta',
      'ie','numCachos','areaFoliar'
    ];
    const mediaCellStyle = { alignment: { horizontal:'center', vertical:'middle' }, border: borda };
    const mediaCalcStyle = {
      alignment: { horizontal:'center', vertical:'middle' }, border: borda,
      fill: { type:'pattern', pattern:'solid', fgColor: { argb:'FFE8F5E9' } }
    };

    let row2 = 2;
    for (let f = 1; f <= TOTAL_FILEIRAS; f++) {
      for (let b = 1; b <= TOTAL_BLOCOS; b++) {
        const plantas = [];
        for (let n = 1; n <= PLANTAS_POR_BLOCO; n++) {
          const r = dados[chave(f, b, n)];
          if (r && !r.ausente) plantas.push(r);
        }
        const media = (campo) => {
          const vals = plantas.map(p => p[campo]).filter(v => v != null && !isNaN(v));
          if (vals.length === 0) return '-';
          return Number((vals.reduce((a, v) => a + v, 0) / vals.length).toFixed(4));
        };

        sheet2.getCell(row2, 1).value = rotulo(f, b);
        sheet2.getCell(row2, 1).style = mediaCellStyle;

        camposMedidas.forEach((campo, i) => {
          const v = media(campo);
          sheet2.getCell(row2, 2 + i).value = v;
          const isCalc = campo === 'spadPlanta' || campo === 'ie' || campo === 'areaFoliar';
          sheet2.getCell(row2, 2 + i).style = isCalc ? mediaCalcStyle : mediaCellStyle;
        });
        row2++;
      }
    }

    row2++;
    const nota2 = sheet2.getCell(row2, 1);
    nota2.value = 'Médias calculadas somente com plantas registradas (ausentes e não registradas são excluídas do cálculo).';
    nota2.font  = { italic: true, size: 9, color: { argb: 'FF555555' } };
    sheet2.mergeCells(row2, 1, row2, 11);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=avaliacao_morfofisiologica.xlsx');
    await workbook.xlsx.write(res);

  } catch (e) {
    console.error('/gerar-excel:', e.message);
    res.status(500).json({ erro: e.message });
  }
});

app.listen(3000, () => console.log('Servidor em http://localhost:3000'));