import XLSX from 'xlsx-js-style';

//* User headers
export function updateUserHeadersStyle(worksheet, user) {
  for (let R = user.s.r; R <= user.e.r; ++R) {
    for (let C = user.s.c; C <= user.e.c; ++C) {
      const cell = worksheet[XLSX.utils.encode_cell({ c: C, r: R })];
      cell.s = {
        font: {
          italic: true,
        },
        fill: {
          fgColor: { rgb: 'F0FFF0' },
        },
      };
    }
  }
}

//* User Info
export function updateUserInfoStyle(worksheet, userInfo) {
  for (let R = userInfo.s.r; R <= userInfo.e.r; ++R) {
    for (let C = userInfo.s.c; C <= userInfo.e.c; ++C) {
      const cellAddress = { c: C, r: R };
      var cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
      if (!cell) {
        worksheet[XLSX.utils.encode_cell(cellAddress)] = {};
        cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
      }
      cell.s = cell.s || {};
      cell.s.fill = {
        fgColor: { rgb: 'FFFAF0' },
      };
    }
  }
}

//* Single Cells
export function updateSingleCellStyle(worksheet) {
  worksheet['A1'].s = {
    font: {
      sz: 24,
      bold: true,
      color: { rgb: '708090' },
    },
  };
  worksheet['A3'].s = {
    font: {
      sz: 12,
      bold: true,
      color: { rgb: '000000' },
    },
    fill: { fgColor: { rgb: '66CDAA' } },
    border: {
      bottom: { style: 'medium', color: { auto: 1 } },
      left: { style: 'medium', color: { auto: 1 } },
      top: { style: 'medium', color: { auto: 1 } },
    },
  };
  worksheet['B3'].s = {
    fill: { fgColor: { rgb: '66CDAA' } },
    border: {
      bottom: { style: 'medium', color: { auto: 1 } },
      right: { style: 'medium', color: { auto: 1 } },
      top: { style: 'medium', color: { auto: 1 } },
    },
  };
  worksheet['A15'].s = {
    font: {
      sz: 12,
      bold: true,
      color: { rgb: '000000' },
    },
    fill: { fgColor: { rgb: '66CDAA' } },
    border: {
      bottom: { style: 'medium', color: { auto: 1 } },
      left: { style: 'medium', color: { auto: 1 } },
      top: { style: 'medium', color: { auto: 1 } },
    },
  };
  worksheet['B15'].s = {
    fill: { fgColor: { rgb: '66CDAA' } },
    border: {
      bottom: { style: 'medium', color: { auto: 1 } },
      right: { style: 'medium', color: { auto: 1 } },
      top: { style: 'medium', color: { auto: 1 } },
    },
  };
}

//* Steps header
export function updateStepsHeaderStyle(worksheet, stepsHeader) {
  for (let R = stepsHeader.s.r; R <= stepsHeader.e.r; ++R) {
    for (let C = stepsHeader.s.c; C <= stepsHeader.e.c; ++C) {
      const cell = worksheet[XLSX.utils.encode_cell({ c: C, r: R })];
      cell.s = {
        font: { bold: true },
        alignment: { horizontal: 'center' },
        fill: {
          fgColor: { rgb: 'D2691E' },
        },
        border: {
          bottom: { style: 'medium', color: { auto: 1 } },
          right: { style: 'medium', color: { auto: 1 } },
        },
      };
    }
  }
}

//* Steps Background Style
export function updateStepsBackgroundStyle(worksheet, steps) {
  for (let rowNum = steps.s.r; rowNum <= steps.e.r; rowNum++) {
    for (let colNum = steps.s.c; colNum <= steps.e.c; colNum++) {
      const cellRef = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
      worksheet[cellRef].s = {
        font: { italic: true },
        alignment: { horizontal: 'center' },
        fill: { fgColor: { rgb: 'FAEBD7' } },
      };
    }
  }
}

//* Actions Background Changes
export function updateActionsBackgroundStyle(worksheet, actions) {
  for (let rowNum = actions.s.r; rowNum <= actions.e.r; rowNum++) {
    for (let colNum = actions.s.c; colNum <= actions.e.c; colNum++) {
      const cellRef = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
      worksheet[cellRef].s = {
        fill: { fgColor: { rgb: 'FFFFE0' } },
        alignment: { horizontal: 'center' },
      };
    }
  }
}

//* Column width changes
export function columnWidthChanges(worksheet) {
  worksheet['!cols'] = [
    { wch: 15 },
    { wch: 35 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 10 },
    { wch: 10 },
    { wch: 20 },
    { wch: 35 },
    { wch: 35 },
  ];
}
