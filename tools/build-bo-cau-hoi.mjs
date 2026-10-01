// Rebuilds the data block of bo-cau-hoi.html from the Word file.
//
//   node tools/build-bo-cau-hoi.mjs ["Bo Cau Hoi - cap nhat.docx"] [bo-cau-hoi.html]
//
// Needs pandoc. If @huggingface/transformers can be resolved from the current
// directory, the script also precomputes the semantic-search vectors
// (multilingual-e5-small); otherwise the page computes them in the browser.
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const docxPath = resolve(process.argv[2] || resolve(root, 'Bo Cau Hoi - cap nhat.docx'));
const htmlPath = resolve(process.argv[3] || resolve(root, 'bo-cau-hoi.html'));
const MODEL = 'Xenova/multilingual-e5-small';

const TOPICS = [
  ['pham-vi', 'Mục tiêu & phạm vi'],
  ['kinh-phi', 'Kinh phí & chi phí'],
  ['cong-nghe', 'Công nghệ & kiến trúc'],
  ['du-lieu', 'Dữ liệu & bản quyền'],
  ['chat-luong', 'Chất lượng & kiểm duyệt'],
  ['do-luong', 'Chỉ tiêu & đo lường'],
  ['ngon-ngu', 'Ngôn ngữ'],
  ['bao-mat', 'An toàn thông tin'],
  ['ca-nhan', 'Dữ liệu cá nhân'],
  ['phap-ly', 'Pháp lý & tuân thủ'],
  ['van-hanh', 'Vận hành & chuyển giao'],
  ['doi-ngu', 'Đội ngũ & tiến độ'],
].map(([id, label]) => ({ id, label }));

// Topics for a section; questions not listed in QUESTION_TOPICS fall back to these.
const SECTION_TOPICS = {
  1: ['pham-vi'], 2: ['pham-vi', 'cong-nghe'], 3: ['cong-nghe'], 4: ['cong-nghe'], 5: ['cong-nghe', 'du-lieu'],
  6: ['du-lieu'], 7: ['ngon-ngu'], 8: ['chat-luong'], 9: ['do-luong'], 10: ['bao-mat'], 11: ['ca-nhan', 'phap-ly'],
  12: ['bao-mat'], 13: ['kinh-phi'], 14: ['van-hanh'], 15: ['doi-ngu'], 16: ['chat-luong', 'van-hanh'],
};
// Keyed by the start of the question (accents and case ignored).
const QUESTION_TOPICS = {
  'du khach da co the dung mien phi chatgpt': ['pham-vi', 'cong-nghe'],
  'ngoai san pham phan mem, dong gop khoa hoc': ['pham-vi', 'do-luong'],
  'cong du lich thong minh gia lai': ['pham-vi', 'cong-nghe'],
  'khoan kinh phi 300 trieu dong cho api': ['kinh-phi', 'cong-nghe'],
  'gia ca, gio mo cua, lich su kien': ['chat-luong', 'du-lieu'],
  'neu chatbot tu van sai': ['chat-luong', 'phap-ly'],
  'cam ket thoi gian phan hoi khong qua 03 giay': ['do-luong', 'cong-nghe'],
  'he thong co phai xac dinh cap do an ninh mang': ['bao-mat', 'phap-ly'],
  'quy trinh xu ly khi xay ra su co': ['bao-mat', 'ca-nhan'],
  'ai chiu trach nhiem bao dam an toan thong tin': ['bao-mat', 'doi-ngu'],
  'he thong thu thap nhung du lieu ca nhan nao': ['ca-nhan'],
  'du lieu vi tri cua du khach': ['ca-nhan'],
  'du lieu dau vao cua co so du lieu du lich': ['ca-nhan', 'du-lieu'],
  'viec su dung api mo hinh ngon ngu lon thuong mai': ['bao-mat', 'ca-nhan'],
  'he thong ngan thong tin ca nhan': ['bao-mat', 'ca-nhan'],
  'lam the nao ngan chatbot bi lam dung': ['bao-mat', 'kinh-phi'],
  'he thong kiem soat the nao viec chatbot': ['chat-luong', 'phap-ly'],
  'he thong dap ung yeu cau minh bach': ['phap-ly'],
  'vi sao de tai chon mo hinh thuong mai': ['cong-nghe', 'ca-nhan'],
  'ha tang dam may du kien su dung': ['kinh-phi', 'cong-nghe'],
  'don vi nao se chi tra': ['kinh-phi', 'van-hanh'],
  'doi lai tinh nhan duoc gi': ['pham-vi', 'kinh-phi'],
  'ket qua cua de tai (ma nguon': ['van-hanh', 'phap-ly'],
  'nhom thuc hien chua co thanh vien chuyen nganh': ['doi-ngu', 'chat-luong'],
  'cac buoc xu ly rui ro chat luong ai': ['chat-luong'],
};

const fold = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
const decode = s => s.replace(/&(#x?[0-9a-f]+|amp|lt|gt|quot|apos|nbsp);/gi, (_, e) =>
  ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' })[e.toLowerCase()] ??
  String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)));
const text = s => decode(s.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').replace(/\s+([,.;:!?)])/g, '$1').replace(/\(\s+/g, '(').trim();
const sentenceCase = s => {
  const t = s.toLocaleLowerCase('vi');
  return (t[0].toLocaleUpperCase('vi') + t.slice(1)).replace(/\b(ai|llm|api|rag|kpi)\b/g, m => m.toUpperCase());
};

// Top-level elements of pandoc's HTML output.
function splitBlocks(src) {
  const out = [], re = /<(\/?)([a-zA-Z0-9]+)\b[^>]*?(\/?)>/g;
  let depth = 0, start = -1, m;
  while ((m = re.exec(src))) {
    const [, close, name, self] = m;
    if (self || /^(br|hr|img|col|meta|link|input|wbr)$/i.test(name)) continue;
    if (!close) { if (depth++ === 0) start = m.index; }
    else if (--depth === 0) out.push(src.slice(start, re.lastIndex));
  }
  return out;
}

// Plain-text lines of one block: one per paragraph, list item or table row.
function blockLines(b) {
  if (/^<(ul|ol)/.test(b)) return [...b.matchAll(/<li>([\s\S]*?)<\/li>/g)].map(m => '• ' + text(m[1]));
  if (/^<table/.test(b)) return [...b.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)]
    .map(m => [...m[1].matchAll(/<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/g)].map(c => text(c[1])).join(' | '));
  if (/^<blockquote/.test(b)) return [...b.matchAll(/<p>([\s\S]*?)<\/p>/g)].map(m => text(m[1]));
  return [text(b)];
}

function cleanHtml(b, sectionNums) {
  b = b.replace(/<colgroup>[\s\S]*?<\/colgroup>/g, '').replace(/\s(style|class|width)="[^"]*"/g, '');
  if (/^<table/.test(b)) {
    const last = [...b.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)]
      .map(r => [...r[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].pop()).filter(Boolean).map(c => text(c[1]));
    const numeric = last.length && last.filter(t => /^(khoảng\s*)?[\d.,]+$/i.test(t)).length >= 0.8 * last.length;
    b = `<div class="tbl${numeric ? ' num-last' : ''}">${b}</div>`;
  }
  // "(mục 5, 8)" points at sections of this document; "mục 13.4 Thuyết minh" does not.
  return b.split(/(<[^>]+>)/).map(part => part.startsWith('<') ? part : part
    .replace(/\[([^\]]*bổ sung[^\]]*)\]/g, '<mark class="todo">[$1]</mark>')
    .replace(/([Mm]ục) (\d+(?:, \d+)*)(?![.,]?\d)(?! Thuyết minh)/g, (all, word, nums) =>
      nums.split(', ').every(n => sectionNums.has(n))
        ? word + ' ' + nums.split(', ').map(n => `<a class="xref" href="#sec-${n}" data-sec="${n}">${n}</a>`).join(', ')
        : all)).join('');
}

function parse(html) {
  const blocks = splitBlocks(html);
  const sectionNums = new Set(blocks.map(text).map(t => (t.match(/^(\d+)\. /) || [])[1]).filter(Boolean));
  const data = { title: '', subtitle: '', updated: '', parts: [], sections: [], topics: TOPICS, items: [] };
  let part = null, sec = null, item = null;
  const finish = () => {
    if (!item) return;
    item.t = item.lines.join('\n');
    item.a = item.html.join('\n');
    delete item.lines; delete item.html;
    item = null;
  };
  for (const b of blocks) {
    const t = text(b);
    if (!data.title && /^<p><strong>/.test(b) && !/^Q\./.test(t)) { data.title = t; continue; }
    if (!data.updated && /^Bản cập nhật ngày/.test(t)) { data.updated = t.replace(/^Bản cập nhật ngày\s*/, ''); continue; }
    if (!data.subtitle && /^<p><em>/.test(b) && !data.parts.length) { data.subtitle = t; continue; }
    // The appendix (questions still waiting for figures) is left out of the page.
    if (/^PHỤ LỤC\b/.test(t)) break;
    let m;
    if ((m = t.match(/^PHẦN ([IVXLC]+)\. (.+)$/))) {
      finish();
      part = { id: 'p' + (data.parts.length + 1), num: 'Phần ' + m[1], title: sentenceCase(m[2]) };
      data.parts.push(part);
      sec = null;
      continue;
    }
    if (/^<p>/.test(b) && !/^<p><(strong|em)>/.test(b) && (m = t.match(/^(\d+)\. (.+)$/))) {
      finish();
      sec = { id: m[1], num: m[1], title: m[2], part: part.id };
      data.sections.push(sec);
      continue;
    }
    if (/^<p><strong>Q\./.test(b)) {
      finish();
      const n = data.items.filter(x => x.sec === sec.id).length + 1;
      const q = t.replace(/^Q\.\s*/, '');
      const key = Object.keys(QUESTION_TOPICS).find(k => fold(q).startsWith(k));
      item = { id: `${sec.num}.${n}`, sec: sec.id, q, topics: key ? QUESTION_TOPICS[key] : SECTION_TOPICS[sec.num], lines: [], html: [] };
      data.items.push(item);
      continue;
    }
    if (!item) { console.warn('Skipped block outside a question:', t.slice(0, 80)); continue; }
    item.lines.push(...blockLines(b).map(l => l.replace(/^A\.\s*/, '')));
    item.html.push(cleanHtml(b.replace(/^<p><strong>A\.<\/strong>\s*/, '<p>'), sectionNums));
  }
  finish();
  const used = new Set(Object.keys(QUESTION_TOPICS).filter(k => data.items.some(it => fold(it.q).startsWith(k))));
  for (const k of Object.keys(QUESTION_TOPICS)) if (!used.has(k)) console.warn('Topic override matched no question:', k);
  return data;
}

// Each answer is split into chunks of a few paragraphs so long answers still
// match on their details. Vector 0 of every item is the question alone.
function chunks(item, maxWords = 90) {
  const lines = item.t.split('\n'), out = [];
  let pos = 0, cur = null;
  for (const line of lines) {
    const words = line.split(/\s+/).length;
    if (cur && cur.words + words > maxWords) { out.push(cur); cur = null; }
    if (!cur) cur = { start: pos, end: pos, words: 0, text: [] };
    cur.text.push(line); cur.words += words; cur.end = pos + line.length;
    pos += line.length + 1;
  }
  if (cur) out.push(cur);
  return out;
}

async function embed(data) {
  let T;
  try {
    const req = createRequire(resolve(process.cwd(), 'noop.js'));
    T = await import(pathToFileURL(req.resolve('@huggingface/transformers')).href);
  } catch {
    console.log('@huggingface/transformers not found from the current directory: vectors will be computed in the browser.');
    return null;
  }
  const extractor = await T.pipeline('feature-extraction', MODEL, { dtype: 'q8' });
  const texts = [], index = [], spans = [];
  for (const it of data.items) {
    const cs = chunks(it);
    index.push([texts.length, cs.length + 1]);
    texts.push('passage: ' + it.q); spans.push(null);
    for (const c of cs) { texts.push('passage: ' + it.q + '\n' + c.text.join('\n')); spans.push([c.start, c.end]); }
  }
  const bytes = [];
  for (let i = 0; i < texts.length; i += 16) {
    const out = await extractor(texts.slice(i, i + 16), { pooling: 'mean', normalize: true });
    const [n, dim] = out.dims;
    for (let r = 0; r < n; r++) {
      const v = out.data.subarray(r * dim, (r + 1) * dim);
      const max = Math.max(...v.map(Math.abs));
      for (const x of v) bytes.push(Math.round(x / max * 127));
    }
    process.stdout.write(`\rEmbedding ${Math.min(i + 16, texts.length)}/${texts.length}`);
  }
  process.stdout.write('\n');
  return { model: MODEL, dim: 384, index, spans, vecs: Buffer.from(Int8Array.from(bytes).buffer).toString('base64') };
}

const html = execFileSync('pandoc', [docxPath, '-t', 'html', '--wrap=none'], { encoding: 'utf8' });
const data = parse(html);
data.emb = await embed(data);

const page = readFileSync(htmlPath, 'utf8');
const re = /(<script id="qa-data" type="application\/json">)[\s\S]*?(<\/script>)/;
if (!re.test(page)) throw new Error('No <script id="qa-data"> block in ' + htmlPath);
const json = JSON.stringify(data).replace(/</g, '\\u003c');
writeFileSync(htmlPath, page.replace(re, (_, a, b) => a + json + b));
console.log(`${data.items.length} questions, ${data.sections.length} sections → ${htmlPath}`);
