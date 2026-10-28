const fs = require('fs');
const path = require('path');

// Simple CSV line parser handling quotes
function parseCsvLine(text) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (inQuotes && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      result.push(cur.trim());
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur.trim());
  return result;
}

const content = fs.readFileSync(path.join(__dirname, 'raw_stock_data.csv'), 'utf8');
const lines = content.split(/\r?\n/);

let section = 'none';
const rawComputers = [];
const rawPrinters = [];
const rawProjectors = [];
const rawIndex = [];

for (let i = 0; i < lines.length; i++) {
  const line = lines[i].trim();
  if (!line || /^,+$/.test(line)) continue;

  if (line.includes('Available Computer Details')) {
    section = 'computers';
    continue;
  } else if (line.startsWith('Index')) {
    section = 'index';
    continue;
  } else if (line.includes('Available Printer Details')) {
    section = 'printers';
    continue;
  } else if (line.includes('Available Projector Details')) {
    section = 'projectors';
    continue;
  }

  if (section === 'computers') {
    if (line.startsWith('Sr. No')) continue;
    const parts = parseCsvLine(line);
    // filter out pure empty parts
    if (parts.some(p => p.length > 0)) {
      rawComputers.push(parts);
    }
  } else if (section === 'index') {
    const parts = parseCsvLine(line);
    if (parts.some(p => p.length > 0)) {
      rawIndex.push(parts);
    }
  } else if (section === 'printers') {
    if (line.startsWith('Sr. No') || line.startsWith(',Laser') || line.startsWith(',InkJet') || line.startsWith(',Thermal') || line.startsWith(',Total') || line.startsWith(',Multifunction')) continue;
    const parts = parseCsvLine(line);
    if (parts.some(p => p.length > 0)) {
      rawPrinters.push(parts);
    }
  } else if (section === 'projectors') {
    if (line.startsWith('Sr. No')) continue;
    const parts = parseCsvLine(line);
    if (parts.some(p => p.length > 0)) {
      rawProjectors.push(parts);
    }
  }
}

console.log('Computers:', rawComputers.length);
console.log('Index:', rawIndex.length);
console.log('Printers:', rawPrinters.length);
console.log('Projectors:', rawProjectors.length);

// Department standardizer function
function normalizeDept(rawDept) {
  if (!rawDept) return 'Unassigned';
  let d = rawDept.trim();
  // Strip trailing/leading spaces, special formatting
  // Standardize common variations
  const lower = d.toLowerCase();
  if (lower === 'bcs' || lower.includes('bcs') && !lower.includes('old') && !lower.includes('stat') && !lower.includes('hod')) return 'BCS';
  if (lower.includes('bcs') && lower.includes('old')) return 'BCS Old PC';
  if (lower.includes('bcs') && lower.includes('hod')) return 'BCS (HOD Office)';
  if (lower.includes('bcs') && lower.includes('stat')) return 'BCS (Stat & Electronics)';
  if (lower === 'language' || lower.includes('language lab')) return 'Language Lab';
  if (lower === 'stat & electronic' || lower === 'stat & electronics') return 'Stat & Electronics';
  if (lower === 'maths' || lower === 'mathematics') return 'Mathematics';
  if (lower.includes('library')) return 'Library';
  if (lower.includes('microbiolology') || lower.includes('microbiology')) return 'Microbiology';
  if (lower.includes('history')) return 'History';
  if (lower === 'srpd') return 'SRPD';
  if (lower.includes('srpd old')) return 'SRPD Old PC';
  if (lower.includes('physics bsc (pg)') || lower.includes('physics pg dept') || lower.includes('pg lab (phy,chem)')) return 'PG Lab (Physics & Chemistry)';
  if (lower === 'physics bsc' || lower.includes('physics bsc') || lower.includes('physics sr wing')) return 'Physics BSc (Sr Wing)';
  if (lower.includes('jr. physics') || lower.includes('physics jr')) return 'Jr. Physics';
  if (lower.includes('bsc cs old')) return 'BSc CS Old PC';
  if (lower.includes('bsc cs (optional)')) return 'BSc CS (Optional)';
  if (lower === 'bsc cs' || lower.includes('bsc cs')) return 'BSc CS';
  if (lower === 'b.sc it + msc  cs' || lower.includes('b.sc it + msc cs')) return 'B.Sc IT + MSc CS';
  if (lower === 'msc cs' || lower === 'msc  cs') return 'MSc CS';
  if (lower === 'msc electronics') return 'MSc Electronics';
  if (lower === 'b.sc it' || lower === 'b.sc it ') return 'B.Sc IT';
  if (lower === 'electronics m.sc') return 'MSc Electronics';
  if (lower === 'electronics' || lower.includes('electronics sr wing') || lower.includes('electronics dept')) return 'Electronics';
  if (lower.includes('geography')) return 'Geography';
  if (lower === 'stat' || lower === 'statistics') return 'Statistics';
  if (lower.includes('website')) return 'Website';
  if (lower.includes('zoology')) return 'Zoology';
  if (lower.includes('botany')) return 'Botany';
  if (lower.includes('principal')) return 'Principal Sir Cabin';
  if (lower.includes('naac')) return 'NAAC';
  if (lower.includes('ncc')) return 'NCC';
  if (lower === 'office') return 'Office';
  if (lower === 'ng office') return 'NG Office';
  if (lower === 'erp') return 'ERP';
  if (lower.includes('girls hostel')) return 'Girls Hostel';
  if (lower === 'hindi') return 'Hindi';
  if (lower === 'english') return 'English';
  if (lower.includes('yogshikshak')) return 'Diploma in Yogshikshak';
  if (lower === 'marathi') return 'Marathi';
  if (lower === 'sro') return 'SRO';
  if (lower.includes('sangli jilha') || lower.includes('rojgar')) return 'Sangli Jilha Rojgar Srujan Kendra';
  if (lower.includes('exam 18')) return 'Exam 18 Room';
  if (lower.includes('economics')) return 'Economics';
  if (lower.includes('sociology')) return 'Sociology';
  if (lower.includes('psychology')) return 'Psychology';
  if (lower.includes('sanskrit')) return 'Sanskrit';
  if (lower.includes('jr. it computer lab') || lower.includes('jr it lab')) return 'Jr. IT Computer Lab';
  if (lower.includes('jr  old pc') || lower.includes('jr old pc')) return 'Jr. Old PC';
  if (lower.includes('chemistry msc') || lower.includes('msc chemistory')) return 'Chemistry MSc';
  if (lower.includes('chemistry bsc') || lower.includes('sr wing chemistry')) return 'Chemistry BSc';
  if (lower.includes('gymkhana')) return 'Gymkhana';
  if (lower.includes('nss')) return 'NSS';
  if (lower.includes('store')) return 'Store Room';
  if (lower.includes('cap junior') || lower.includes('cap 1st 2nd year')) return 'CAP Junior / Exam Section';
  if (lower.includes('biotechnology')) return 'Biotechnology';
  if (lower.includes('velankar hall')) return 'Velankar Hall';
  if (lower.includes('class room no 8')) return 'Class Room No 8 (Language Lab)';
  if (lower.includes('class room no 9')) return 'Class Room No 9 (BSc IT AS 7)';
  if (lower.includes('class room no 10')) return 'Class Room No 10 (MSc CS)';
  if (lower.includes('class room no 20')) return 'Class Room No 20 (BSc IT AS 6)';

  return d;
}

const deptMap = {};

function getDept(dName) {
  const norm = normalizeDept(dName);
  if (!deptMap[norm]) {
    deptMap[norm] = {
      department: norm,
      computers: [],
      printers: [],
      projectors: [],
      stats: {
        totalEquipment: 0,
        computers: { total: 0, working: 0, notWorking: 0 },
        printers: { total: 0, working: 0, notWorking: 0 },
        projectors: { total: 0, working: 0, notWorking: 0 }
      }
    };
  }
  return deptMap[norm];
}

// Process computers
rawComputers.forEach((row, idx) => {
  // Columns: Sr. No, Dead Stock Register Number, Department, Make, Type, Monitor, Motherboard, C.P.U., RAM, HDD, SSD, O. S., Keyboard, Mouse, Status, Remark...
  let srNo = row[0] || '';
  let deadStock = row[1] || '';
  let dept = row[2] || '';
  let make = row[3] || '';
  let type = row[4] || 'Desktop';
  let monitor = row[5] || '';
  let motherboard = row[6] || '';
  let cpu = row[7] || '';
  let ram = row[8] || '';
  let hdd = row[9] || '';
  let ssd = row[10] || '';
  let os = row[11] || '';
  let keyboard = row[12] || '';
  let mouse = row[13] || '';
  let status = row[14] || 'Working';
  let remark = row[15] || '';

  // Clean status
  const sLower = status.toLowerCase();
  let isWorking = true;
  let faultIssue = null;

  if (sLower.includes('not working') || sLower.includes('issue') || sLower.includes('no ram') || sLower.includes('smps')) {
    isWorking = false;
    faultIssue = status;
    status = 'Not Working';
  } else if (sLower.includes('working')) {
    status = 'Working';
  } else if (!status) {
    status = 'Working';
  }

  if (dept) {
    const dObj = getDept(dept);
    const pcId = `WCS-PC-${srNo || (idx + 1)}`;
    const compRecord = {
      id: pcId,
      srNo: srNo || (idx + 1),
      deadStock,
      department: dObj.department,
      originalDept: dept,
      make,
      type: 'Computer',
      deviceType: type,
      monitor,
      motherboard,
      cpu,
      ram,
      hdd,
      ssd,
      os,
      keyboard,
      mouse,
      status,
      isWorking,
      faultIssue,
      remark
    };
    dObj.computers.push(compRecord);
    dObj.stats.computers.total++;
    if (isWorking) dObj.stats.computers.working++;
    else dObj.stats.computers.notWorking++;
    dObj.stats.totalEquipment++;
  }
});

// Process printers
rawPrinters.forEach((row, idx) => {
  // Sr. No, Dept., Printer Asset number, Make, Type, PORT, Remark
  let srNo = row[0] || (idx + 1);
  let dept = row[1] || '';
  let assetNum = row[2] || '';
  let make = row[3] || '';
  let pType = row[4] || 'Laser';
  let port = row[5] || 'USB';
  let remark = row[6] || '';

  let isWorking = !remark.toLowerCase().includes('not working');
  let status = isWorking ? 'Working' : 'Not Working';

  if (dept) {
    const dObj = getDept(dept);
    const prRecord = {
      id: `WCS-PRN-${srNo}`,
      srNo,
      department: dObj.department,
      originalDept: dept,
      assetNumber: assetNum,
      make,
      type: 'Printer',
      printerType: pType,
      port,
      status,
      isWorking,
      remark
    };
    dObj.printers.push(prRecord);
    dObj.stats.printers.total++;
    if (isWorking) dObj.stats.printers.working++;
    else dObj.stats.printers.notWorking++;
    dObj.stats.totalEquipment++;
  }
});

// Process projectors
rawProjectors.forEach((row, idx) => {
  // Sr. No, Dept., Serial Number number, Make, Remark
  let srNo = row[0] || (idx + 1);
  let dept = row[1] || '';
  let serialNum = row[2] || '';
  let make = row[3] || '';
  let remark = row[4] || '';

  let isWorking = !remark.toLowerCase().includes('not working');
  let status = isWorking ? 'Working' : 'Not Working';

  if (dept) {
    const dObj = getDept(dept);
    const pjRecord = {
      id: `WCS-PRJ-${srNo}`,
      srNo,
      department: dObj.department,
      originalDept: dept,
      serialNumber: serialNum,
      make,
      type: 'Projector',
      status,
      isWorking,
      remark
    };
    dObj.projectors.push(pjRecord);
    dObj.stats.projectors.total++;
    if (isWorking) dObj.stats.projectors.working++;
    else dObj.stats.projectors.notWorking++;
    dObj.stats.totalEquipment++;
  }
});

const sortedDepts = Object.keys(deptMap).sort();
console.log(`Total Unique Departments: ${sortedDepts.length}`);

// Generate summary report
const summaryReport = sortedDepts.map(dName => {
  const d = deptMap[dName];
  return {
    department: dName,
    computers: d.stats.computers,
    printers: d.stats.printers,
    projectors: d.stats.projectors,
    totalEquipment: d.stats.totalEquipment
  };
});

fs.writeFileSync('department_split_summary.json', JSON.stringify(summaryReport, null, 2));
fs.writeFileSync('department_complete_data.json', JSON.stringify(deptMap, null, 2));
console.log('Saved department_split_summary.json and department_complete_data.json');

