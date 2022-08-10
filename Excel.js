const ExcelJS = require('exceljs');
import saveAs from 'file-saver';

export function excelExport({
  title,
  filename,
  headerName = [],
  exportData = [],
  autoWidth = true,
  autoWrap = true,
  exportWidth = [],
  sheetName = "sheet"
} = {}) {
  const owner = "XXXXXX有限公司"
  const workbook = new ExcelJS.Workbook();

  //代码中出现的&开头字符对应变量
  const pageHeader = "&R&KDBDBDB&9&\"等线\"" + title
  const pageFooter = "&L&KDBDBDB&9&\"等线\"" + owner + "&C&KDBDBDB&9&\"等线\"&P/&N&R&KDBDBDB&9&\"等线\"未经允许 严禁传播"

  //创建工作表
  const workSheet = workbook.addWorksheet(sheetName);
  workSheet.headerFooter.firstHeader = pageHeader
  workSheet.headerFooter.firstFooter = pageFooter
  workSheet.headerFooter.oddHeader = pageHeader
  workSheet.headerFooter.oddFooter = pageFooter
  workSheet.headerFooter.evenHeader = pageHeader
  workSheet.headerFooter.evenFooter = pageFooter

  //列数
  const headerLength = headerName.length
  //行数
  const dataLength = exportData.length

  //最大列
  const maxColumns = createCol(headerLength - 1)

  //标题行单元格范围
  const headerLine = "A1:" + maxColumns + "1"
  //合并单元格
  workSheet.mergeCells(headerLine);
  //插入标题
  workSheet.getCell(1, 1).value = title + "报价";
  //标题格式
  workSheet.getCell('A1').font = {
    name: '等线',
    family: 4,
    size: 12,
    underline: false,
    bold: true
  };
  workSheet.getCell('A1').alignment = {
    vertical: 'middle', horizontal: 'center'
  };

  if (autoWidth) {
    //设置worksheet每列的最大宽度，是否自动换行
    const colWidth = exportData.map(row => row.map(val => {
      var width = 10
      var wrapText = false
      if (val != null) {
        width = val.toString().length
        if (val.toString().charCodeAt(0) > 255) {
          width = width * 2
        }
        if (autoWrap && width > 99) {
          wrapText = true
          width = 100
        }
      }

      return {
        'width': width,
        'wrapText': wrapText
      };
    }))

    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['width'] < colWidth[i][j]['width']) {
          result[j]['width'] = colWidth[i][j]['width'];
        }
        if (colWidth[i][j]['wrapText']) {
          result[j]['wrapText'] = colWidth[i][j]['wrapText']
        }
      }
    }

    exportWidth = result
  }

  //设置列标题格式
  for (let i = 1; i < headerLength + 1; i++) {
    let value = headerName[i - 1];
    let width = exportWidth[i - 1].width;

    //设置列宽
    const dobCol = workSheet.getColumn(i);
    dobCol.width = width

    //设置标题和标题单元格格式
    const cell = workSheet.getCell(3, i)
    if (value == "数量" || value == "单位") {
      cell.alignment = {
        vertical: 'middle', horizontal: 'center'
      };
    }
    cell.value = value
    cell.font = {
      name: '等线',
      family: 4,
      size: 11,
      bold: true
    };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }

  //插入数据
  for (let i = 4; i < dataLength + 4; i++) {
    for (let j = 0; j < headerLength; j++) {
      workSheet.getCell(i, j + 1).value = exportData[i - 4][j]
      var font = {
        name: '等线',
        family: 4,
        size: 11,
      };
      if (exportData[i - 4][1] == "小计" || exportData[i - 4][1] == "合计") {
        font = {
          name: '等线',
          family: 4,
          size: 11,
          bold: true
        }
      }
      if (headerName[j] == "数量" || headerName[j] == "单位") {
        workSheet.getCell(i, j + 1).alignment = {
          vertical: 'middle', horizontal: 'center'
        };
      }
      workSheet.getCell(i, j + 1).font = font
      if (autoWrap && exportWidth[j].wrapText) {
        workSheet.getCell(i, j + 1).alignment = { wrapText: true }
      }
      workSheet.getCell(i, j + 1).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  //尾注:公司名和日期
  var a = dataLength + 6
  var b = dataLength + 7
  const endnoteStartColumn = createCol(headerLength - 4) + a
  const endnoteEndColumn = maxColumns + a
  const dateStartColumn = createCol(headerLength - 4) + b
  const dateEndColumn = maxColumns + b

  const aCell = workSheet.getCell(a, headerLength - 1)
  const bCell = workSheet.getCell(b, headerLength - 1)

  const aa = endnoteStartColumn + ':' + endnoteEndColumn
  const bb = dateStartColumn + ':' + dateEndColumn

  //合并单元格
  workSheet.mergeCells(aa);
  workSheet.mergeCells(bb);

  //插入尾注
  aCell.value = owner;
  bCell.value = date();
  aCell.font = bCell.font = {
    name: '等线',
    family: 4,
    size: 10,
  };
  aCell.alignment = bCell.alignment = {
    vertical: 'middle', horizontal: 'center'
  };

  //导出
  workbook.xlsx.writeBuffer().then(function (buffer) {
    saveAs(new Blob([buffer], {
      type: 'application/octet-stream'
    }), filename + '.' + 'xlsx');
  });
}

/**
 * 数字转字母（按excel形成）
 * createCol(26)=>AA
 * createCol(25)=>Z
 * @param {*} n
 * @returns
 */
function createCol(n) {
  const ordA = 'A'.charCodeAt(0)
  const ordZ = 'Z'.charCodeAt(0)
  const len = ordZ - ordA + 1
  let str = ""
  while (n >= 0) {
    str = String.fromCharCode(n % len + ordA) + str
    n = Math.floor(n / len) - 1
  }
  return str
}

function date() {
  var nowDate = new Date();
  var year = nowDate.getFullYear();
  var month = nowDate.getMonth() + 1 < 10 ? "0" + (nowDate.getMonth() + 1)
    : nowDate.getMonth() + 1;
  var day = nowDate.getDate() < 10 ? "0" + nowDate.getDate() : nowDate
    .getDate();
  var dateStr = year + "年" + month + "月" + day + "日";
  return dateStr
}
