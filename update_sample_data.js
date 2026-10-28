const fs = require('fs');
const path = require('path');

const deptMap = JSON.parse(fs.readFileSync('department_complete_data.json', 'utf8'));
const summary = JSON.parse(fs.readFileSync('department_split_summary.json', 'utf8'));

const departmentList = Object.keys(deptMap).sort();

const allAssets = [];

departmentList.forEach(dName => {
  const d = deptMap[dName];

  // Computers
  d.computers.forEach(c => {
    let specList = [];
    if (c.cpu) specList.push(c.cpu);
    if (c.ram) specList.push(c.ram + ' RAM');
    if (c.hdd) specList.push(c.hdd + ' HDD');
    if (c.ssd) specList.push(c.ssd + ' SSD');
    if (c.os) specList.push(c.os);
    if (c.motherboard) specList.push('MB: ' + c.motherboard);

    const asset = {
      id: c.id,
      type: "Computer",
      name: `${c.make || 'PC'} ${c.cpu ? c.cpu : 'Desktop'}`.trim(),
      department: c.department,
      location: `${c.department} Lab / Room`,
      assignedUser: `${c.department} Faculty / Students`,
      status: c.status,
      serialNumber: c.deadStock || `DS-${c.srNo}`,
      purchaseDate: "2025-01-18",
      specs: specList.join(', ') || 'Standard PC Configuration',
      ipAddress: `192.168.10.${(parseInt(c.srNo) || 1) % 254 + 1}`,
      faultDetails: !c.isWorking ? {
        issue: c.faultIssue || c.remark || 'Hardware fault reported',
        reportedDate: "2025-11-21",
        reportedBy: `${c.department} Lab In-charge`,
        priority: "High",
        technician: "Hardware Maintenance Team"
      } : null,
      excelAttachments: [
        {
          id: `att-${c.id.toLowerCase()}-1`,
          fileName: `${c.id}_Spec_Audit.xlsx`,
          fileSize: "14.2 KB",
          uploadDate: "2025-11-21",
          notes: `Stock Audit details for ${c.id} (${c.department})`,
          sheetData: [
            ["Attribute", "Specification / Value"],
            ["Department", c.department],
            ["Make / Brand", c.make || "N/A"],
            ["Type", c.deviceType || "Desktop"],
            ["Processor (CPU)", c.cpu || "N/A"],
            ["Motherboard", c.motherboard || "N/A"],
            ["RAM", c.ram || "N/A"],
            ["HDD", c.hdd || "N/A"],
            ["SSD", c.ssd || "N/A"],
            ["Operating System", c.os || "N/A"],
            ["Monitor", c.monitor || "N/A"],
            ["Keyboard / Mouse", `${c.keyboard} / ${c.mouse}`],
            ["Operational Status", c.status],
            ["Remarks / Issues", c.faultIssue || c.remark || "None"]
          ]
        }
      ]
    };
    allAssets.push(asset);
  });

  // Printers
  d.printers.forEach(p => {
    const asset = {
      id: p.id,
      type: "Printer",
      name: `${p.make || 'Printer'} (${p.printerType || 'Laser'})`,
      department: p.department,
      location: `${p.department} Office / Desk`,
      assignedUser: `${p.department} Staff`,
      status: p.status,
      serialNumber: p.assetNumber || `PRN-AST-${p.srNo}`,
      purchaseDate: "2025-11-21",
      specs: `Type: ${p.printerType || 'Laser'}, Port: ${p.port || 'USB'}`,
      ipAddress: "",
      faultDetails: !p.isWorking ? {
        issue: p.remark || 'Printer Malfunction',
        reportedDate: "2025-11-21",
        reportedBy: `${p.department} In-charge`,
        priority: "Medium",
        technician: "Printer Service Vendor"
      } : null,
      excelAttachments: []
    };
    allAssets.push(asset);
  });

  // Projectors
  d.projectors.forEach(j => {
    const asset = {
      id: j.id,
      type: "Projector",
      name: `${j.make || 'Projector'}`,
      department: j.department,
      location: `${j.department}`,
      assignedUser: `${j.department} Lecture Hall`,
      status: j.status,
      serialNumber: j.serialNumber || `PRJ-SN-${j.srNo}`,
      purchaseDate: "2026-07-11",
      specs: `Model: ${j.make || 'DLP Projector'}, Remarks: ${j.remark || 'Installed'}`,
      ipAddress: "",
      faultDetails: !j.isWorking ? {
        issue: j.remark || 'Display / Lamp issue',
        reportedDate: "2026-07-11",
        reportedBy: `${j.department} Dept`,
        priority: "High",
        technician: "AV Support Team"
      } : null,
      excelAttachments: []
    };
    allAssets.push(asset);
  });
});

const sampleDataFileContent = `/**
 * Deccan Education Society's Willingdon College, Sangli
 * Live Inventory Seed Data: Computers, Printers, Projectors
 * Split across 60 Departments (Stock Audit Data)
 */

const INITIAL_DEPARTMENTS = ${JSON.stringify(departmentList, null, 2)};

const INITIAL_ASSETS = ${JSON.stringify(allAssets, null, 2)};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { INITIAL_DEPARTMENTS, INITIAL_ASSETS };
}
`;

fs.writeFileSync('sample_data.js', sampleDataFileContent, 'utf8');
console.log(`Updated sample_data.js with ${departmentList.length} departments and ${allAssets.length} total equipment assets!`);

