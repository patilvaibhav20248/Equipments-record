const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

try { fs.unlinkSync(path.join(__dirname, 'test.xlsx')); } catch(e){}

const deptMap = JSON.parse(fs.readFileSync('department_complete_data.json', 'utf8'));
const summary = JSON.parse(fs.readFileSync('department_split_summary.json', 'utf8'));

// Create folder for individual department files
const splitDir = path.join(__dirname, 'departments_split');
if (!fs.existsSync(splitDir)) {
  fs.mkdirSync(splitDir, { recursive: true });
}

// Helper to save workbook with fs
function saveWorkbook(wb, destPath) {
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  fs.writeFileSync(destPath, buf);
}

// 1. Create a Master Excel Workbook
const masterWb = XLSX.utils.book_new();

// Sheet 1: Summary Index
const summaryRows = [
  ['Deccan Education Society\'s Willingdon College, Sangli'],
  ['Department Equipment & Hardware Inventory Summary - Stock Audit (21-11-2025)'],
  [],
  ['Sr No', 'Department Name', 'PC Total', 'PC Working', 'PC Faulty / Not Working', 'Printers', 'Projectors', 'Total Equipment', 'Operational %']
];

summary.forEach((s, idx) => {
  const pcHealth = s.computers.total > 0 ? Math.round((s.computers.working / s.computers.total) * 100) + '%' : '100%';
  summaryRows.push([
    idx + 1,
    s.department,
    s.computers.total,
    s.computers.working,
    s.computers.notWorking,
    s.printers.total,
    s.projectors.total,
    s.totalEquipment,
    pcHealth
  ]);
});

// Add totals row
const totalPcs = summary.reduce((a, b) => a + b.computers.total, 0);
const totalPcWork = summary.reduce((a, b) => a + b.computers.working, 0);
const totalPcFault = summary.reduce((a, b) => a + b.computers.notWorking, 0);
const totalPrinters = summary.reduce((a, b) => a + b.printers.total, 0);
const totalProjectors = summary.reduce((a, b) => a + b.projectors.total, 0);
const totalEquip = summary.reduce((a, b) => a + b.totalEquipment, 0);

summaryRows.push([]);
summaryRows.push([
  'TOTAL',
  'All Departments',
  totalPcs,
  totalPcWork,
  totalPcFault,
  totalPrinters,
  totalProjectors,
  totalEquip,
  Math.round((totalPcWork / totalPcs) * 100) + '%'
]);

const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
XLSX.utils.book_append_sheet(masterWb, wsSummary, 'Summary Index');

// Sheet 2: Master All Equipment
const allEquipmentRows = [
  ['Asset ID', 'Type', 'Department', 'Make', 'Specs / CPU', 'RAM', 'HDD', 'SSD', 'O.S.', 'Monitor / Serial', 'Peripherals (Kbd/Mouse/Port)', 'Status', 'Fault / Remark']
];

Object.keys(deptMap).sort().forEach(dName => {
  const d = deptMap[dName];
  d.computers.forEach(c => {
    allEquipmentRows.push([
      c.id,
      'Computer',
      c.department,
      c.make,
      c.cpu + (c.motherboard ? ' | MB: ' + c.motherboard : ''),
      c.ram,
      c.hdd,
      c.ssd,
      c.os,
      c.monitor || 'Desktop',
      `Kbd: ${c.keyboard} | Mouse: ${c.mouse}`,
      c.status,
      c.faultIssue || c.remark || ''
    ]);
  });

  d.printers.forEach(p => {
    allEquipmentRows.push([
      p.id,
      'Printer',
      p.department,
      p.make,
      p.printerType,
      '',
      '',
      '',
      '',
      p.assetNumber || 'N/A',
      `Port: ${p.port}`,
      p.status,
      p.remark || ''
    ]);
  });

  d.projectors.forEach(j => {
    allEquipmentRows.push([
      j.id,
      'Projector',
      j.department,
      j.make,
      'Projector',
      '',
      '',
      '',
      '',
      j.serialNumber || 'N/A',
      '',
      j.status,
      j.remark || ''
    ]);
  });
});

const wsAll = XLSX.utils.aoa_to_sheet(allEquipmentRows);
XLSX.utils.book_append_sheet(masterWb, wsAll, 'All Equipment');

let deptFilesCount = 0;
// Generate separate workbooks/CSVs for each department
Object.keys(deptMap).sort().forEach(dName => {
  const d = deptMap[dName];
  if (d.stats.totalEquipment === 0) return;

  const safeFileName = dName.replace(/[\/\\?%*:|"<>]/g, '_');
  const deptWb = XLSX.utils.book_new();

  const deptRows = [
    [`Willingdon College, Sangli - Department Inventory: ${dName}`],
    [`Total Assets: ${d.stats.totalEquipment} | Computers: ${d.stats.computers.total} (Working: ${d.stats.computers.working}, Faulty: ${d.stats.computers.notWorking}) | Printers: ${d.stats.printers.total} | Projectors: ${d.stats.projectors.total}`],
    [],
    ['Asset ID', 'Type', 'Make', 'CPU / Model', 'RAM', 'HDD', 'SSD', 'O.S.', 'Monitor / Asset No', 'Keyboard / Mouse / Port', 'Status', 'Fault / Remark']
  ];

  d.computers.forEach(c => {
    deptRows.push([
      c.id,
      'Computer',
      c.make,
      c.cpu + (c.motherboard ? ' | MB: ' + c.motherboard : ''),
      c.ram,
      c.hdd,
      c.ssd,
      c.os,
      c.monitor || 'Desktop',
      `Kbd: ${c.keyboard} | Mouse: ${c.mouse}`,
      c.status,
      c.faultIssue || c.remark || ''
    ]);
  });

  d.printers.forEach(p => {
    deptRows.push([
      p.id,
      'Printer',
      p.make,
      p.printerType,
      '',
      '',
      '',
      '',
      p.assetNumber || 'N/A',
      `Port: ${p.port}`,
      p.status,
      p.remark || ''
    ]);
  });

  d.projectors.forEach(j => {
    deptRows.push([
      j.id,
      'Projector',
      j.make,
      'Projector',
      '',
      '',
      '',
      '',
      j.serialNumber || 'N/A',
      '',
      j.status,
      j.remark || ''
    ]);
  });

  // Excel sheet names cannot exceed 31 chars and cannot have: \ / ? * : [ ]
  let sheetTitle = dName.replace(/[\/\\?%*:|"<>\[\]]/g, ' ').trim().substring(0, 31);
  const wsDept = XLSX.utils.aoa_to_sheet(deptRows);
  XLSX.utils.book_append_sheet(deptWb, wsDept, sheetTitle);

  // Write dept workbook
  saveWorkbook(deptWb, path.join(splitDir, `${safeFileName}.xlsx`));
  deptFilesCount++;

  // Also append major departments to master workbook if total equipment >= 5 and sheet names < 25
  if (masterWb.SheetNames.length < 25 && d.stats.totalEquipment >= 8) {
    XLSX.utils.book_append_sheet(masterWb, wsDept, sheetTitle);
  }
});

// Save Master Excel file
const masterPath = path.join(__dirname, 'Willingdon_College_Inventory_Split_By_Department.xlsx');
saveWorkbook(masterWb, masterPath);
console.log('Successfully written Master Excel file:', masterPath);
console.log(`Successfully generated ${deptFilesCount} individual department Excel files in:`, splitDir);
