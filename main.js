const {app, BrowserWindow, ipcMain, dialog} = require('electron');
const path = require('path');
const xlsx = require('xlsx');
const xlsxs = require('xlsx-style');

function createWindow () {
  const mainWindow = new BrowserWindow({
    width: 500,
    height: 500,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: true
    }
  });
  mainWindow.loadFile('index.html');
};

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
});

app.on('window-all-closed', () => {
  app.quit()
});

ipcMain.on('asynchronous-message', (event, arg) => {
  function selectFile(prop, typ) {
    dialog.showOpenDialog({ properties: [prop] })
    .then(x => {
      const file = x.filePaths[0];
      if (file)
        event.reply('asynchronous-reply', {typ, file});
    })
    .catch(e => console.log(e));
  }

  switch (arg.typ) {
    case 1:
    case 2:
      selectFile('openFile', arg.typ);
      break;
    case 3:
      selectFile('openDirectory', arg.typ);
      break;
    case 4:
      const { file1, file2, dir } = arg;

      const records = getRows(file1).map(obj => new Record(obj));
      const recordsMap = groupBy(records, r => r.id);

      const students = getRows(file2).map(obj => {
        const s = new Student(obj);
        const records = recordsMap.get(s.id);
        s.readRecords(records ? records : []);
        return s;
      });
      const perDept = groupBy(students, s => s.dept);
      
      const date = students.map(s => s.date).find(d => d);
      const y = Number(date.substring(0, 4));
      const m = Number(date.substring(5, 7)) - 1;
      const d = Number(date.substring(8, 10));
      const dateObj = new Date();
      dateObj.setFullYear(y, m, d);
      const dateStr = `${date}(${dayName(dateObj)})`;
      for (const [dept, students] of perDept) {
        function getAbsent(desc, reason) {
          return [
            desc,
            students.filter(s => s.isAbsent && s.absenceReason === reason).length
          ];
        }
        const title1 = '복무현황';
        const data1 = [
          [`${dateStr} ${dept} ${title1}`],
          [],
          ['복무구분', '인원(명)'],
          ['정상출근', students.filter(s => !s.isLate && !s.isAbsent).length],
          ['지각', students.filter(s => s.isLate).length],
          getAbsent('국내출장', '국내출장'),
          getAbsent('국내파견', '국내파견'),
          getAbsent('해외출장', '해외출장'),
          getAbsent('4주군사훈련', '4주훈련소교육'),
          getAbsent('연차휴가', '휴가(연가)_1일이상'),
          getAbsent('병가', '휴가(병가)'),
          getAbsent('경조사휴가', '휴가(경조사)'),
          getAbsent('무단결근', ''),
          ['합계', students.length]
        ];
      
        const title2 = '지각자현황';
        const header = [
          [`${dateStr} ${dept} ${title2}`],
          [],
          ['학과', '학번', '성명', '지각패널티', '누적복무연장']
        ];
        const info = students
          .filter(s => s.isLate)
          .sort((a, b) => a.id - b.id)
          .map(s => s.info());
        const data2 = header.concat(info);

        const ws1 = arrayToSheet(data1, true);
        const ws2 = arrayToSheet(data2);
        const wb = sheetsToBook([[title1, ws1], [title2, ws2]]);
        xlsxs.writeFile(wb, path.join(dir, `${dateStr} ${dept}.xlsx`), { bookType: 'xlsx' });
      }
      break;
  }
});

class Record {
  constructor(obj) {
    this.id = Number(obj['학번/사번']);
    this.time = obj['출근일시'].substring(11);
    if (this.time)
      this.date = obj['출근일시'].substring(0, 10);
    this.etc = obj['비고'];
  }
}

const absenceReason = [
  '4주훈련소교육',
  '국내출장',
  '국내파견',
  '해외출장',
  '휴가(경조사)',
  '휴가(병가)',
  '휴가(연가)_1일이상'
];

function timeToNum(time) {
  const h = Number(time.substring(0, 2));
  const m = Number(time.substring(3));
  return h * 60 + m;
}

function numToTime(num) {
  const h = Math.floor(num / 60);
  const m = num - h * 60;
  function twoDigits(s) {
    return (s < 10) ? `0${s}` : s.toString();
  }
  return `${twoDigits(h)}:${twoDigits(m)}`;
}

function dayName(date) {
  switch (date.getDay()) {
    case 0: return '일';
    case 1: return '월';
    case 2: return '화';
    case 3: return '수';
    case 4: return '목';
    case 5: return '금';
    case 6: return '토';
  }
}

class Student {
  constructor(obj, records) {
    this.dept = obj['학과/부서'];
    this.id = Number(obj['학번/사번']);
    this.name = obj['성명'];
    this.days = obj['패널티일'];
    this.hours = obj['패널티시간'];
  }

  readRecords(records) {
    this.date = records.map(o => o.date).find(d => d);
    this.time = records.map(o => o.time).join('');
    this.etc = records.map(o => o.etc).filter(s => s);
    this.limit = this.etc.map(s => {
      switch (s) {
        case '출근시간변경(11:00)': return '11:00';
        case '출근시간변경(11:30)': return '11:30';
        case '출근시간변경(12:00)': return '12:00';
        case '출근시간변경(12:30)': return '12:30';
        case '출근시간변경(13:00)': return '13:00';
        case '출근시간변경(13:30)': return '13:30';
        case '출근시간변경(14:00)': return '14:00';
        case '휴가(연가)_오전반차': return '14:00';
      }
      return '';
    }).join('');
    if (!this.limit) this.limit = '10:30';
    this.isAbsent = this.time.length === 0;
    this.absenceReason = this.etc.find(s => absenceReason.includes(s));
    if (!this.absenceReason) this.absenceReason = '';
    if (!this.isAbsent) {
      this.delta = timeToNum(this.time) - timeToNum(this.limit);
      this.isLate = this.delta > 0;
    } else {
      this.delta = 0;
      this.isLate = false;
    }
  }

  info() {
    return [
      this.dept,
      this.id,
      this.name,
      numToTime(this.delta),
      `${this.days}일 ${this.hours}시간`
    ];
  }
}

function groupBy(arr, f) {
  const map = new Map();
  for (const elem of arr) {
    const key = f(elem);
    let vs = map.get(key);
    if (!vs) vs = [];
    vs.push(elem);
    map.set(key, vs);
  }
  return map;
}

function getRows(fn) {
  const wb = xlsx.readFile(fn);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return xlsx.utils.sheet_to_json(ws);
}

function widthOf(str) {
  let w = 0;
  for (const i in str) {
    let code = str.charCodeAt(i);
    w += (0xac00 <= code && code <= 0xd7af) ? 2 : 1;
  }
  return w;
}

function arrayToSheet(arr, footer) {
  const ws = {};
  const R = arr.length;
  const C = arr.map(a => a.length).reduce((a, b) => (a > b) ? a : b, 0);
  const range = { s: { c: 0, r: 0 }, e: { c: C - 1, r: R - 1 } };
  const widths = Array(C);
  widths.fill(10);

  for (let r = 0; r < R; r++) {
    for (let c = 0; c < arr[r].length; c++) {
      const cell = { v: arr[r][c] };
      cell.t = (typeof(cell.v) === 'number') ? 'n' : 's';
      if (r === 0)
        cell.s = { font: { bold: true } };
      else if (r === 2 || (r === arr.length - 1 && footer))
        cell.s = { font: { bold: true }, fill: { patternType: 'solid', fgColor: { rgb: "FFD0D0D0" } } };
      else
        cell.s = {};
      if (r !== 0) {
        cell.s.border = {
          top: { style: "thin", color: { rgb: "FF000000" } },
          bottom: { style: "thin", color: { rgb: "FF000000" } },
          left: { style: "thin", color: { rgb: "FF000000" } },
          right: { style: "thin", color: { rgb: "FF000000" } }
        };
        const w = widthOf(cell.v.toString());
        if (w > widths[c]) widths[c] = w;
      }
      const ref = xlsxs.utils.encode_cell({ r, c });
      ws[ref] = cell;
    }
  }
  ws['!ref'] = xlsxs.utils.encode_range(range);
  ws['!cols'] = widths.map(w => { return { wch: w }; });
  ws['!merges'] = [ { s: { c: 0, r: 0 }, e: { c: 4, r: 0 } } ];
  return ws;
}

function sheetsToBook(sheets) {
  const wb = { SheetNames: [], Sheets: {} };
  for (const [name, sheet] of sheets) {
    wb.SheetNames.push(name);
    wb.Sheets[name] = sheet;
  }
  return wb;
}
