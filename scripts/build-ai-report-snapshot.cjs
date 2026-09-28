const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');
const inputPaths = process.argv.slice(2);

if (inputPaths.length !== 4) {
  console.error('用法: node scripts/build-ai-report-snapshot.cjs 第一周会话.xlsx 第二周会话.xlsx 第一周主线索.xlsx 第二周主线索.xlsx');
  process.exit(1);
}

global.window = global;
global.document = { readyState: 'loading', addEventListener() {} };
global.localStorage = { getItem() { return null; }, setItem() {}, removeItem() {} };
global.XLSX = require(path.join(repoRoot, 'assets', 'xlsx.full.min.js'));

vm.runInThisContext(fs.readFileSync(path.join(repoRoot, 'assets', 'ai-report-refresh.js'), 'utf8'));

function readRows(filePath) {
  const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
}

const files = inputPaths.map((filePath) => ({ name: path.basename(filePath) }));
const rows = inputPaths.map(readRows);
const snapshot = __AIReportRefresh.buildSnapshot(files, rows[0], rows[1], rows[2], rows[3]);
const outputPath = path.join(repoRoot, 'data', 'ai-report-snapshot.json');

fs.writeFileSync(outputPath, `${JSON.stringify(snapshot, null, 2)}\n`, 'utf8');
console.log(`已生成 ${outputPath}`);
console.log(`统计周期：${new Date(snapshot.first.period.start).toISOString().slice(0, 10)} 至 ${new Date(snapshot.second.period.end).toISOString().slice(0, 10)}`);
