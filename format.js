const path = require('path')
const fs = require('fs');
const xl = require('excel4node');
const readXlsxFile = require('read-excel-file/node')

// 贵阳市 六盘水市  铜仁市 黔南布依族苗族自治州  遵义市  安顺市  毕节市  黔西南布依族苗族自治州  黔东南苗族侗族自治州
const origin = '黔东南苗族侗族自治州'
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');
ws.cell(1, 1).string('商品编码')
ws.cell(1, 2).string('商品名称')
ws.cell(1, 3).string('商品成本')
ws.cell(1, 4).string('市场价')
ws.cell(1, 5).string('产地')


const genXlsx = (ws, arr = []) => {
  const len = arr.length
  for (let i = 0; i < len; i++) {
    ws.cell(i + 2, 2).string(String(arr[i].name))
    ws.cell(i + 2, 3).number(Number(arr[i].cost))
    ws.cell(i + 2, 4).number(Number(arr[i].sell))
    ws.cell(i + 2, 5).string(String(arr[i].origin))
  }

  wb.write(path.join(__dirname, `./output/${origin}.xlsx`), function(err, stats) {
    if (err) {
      console.error(err);
    } else {
      console.log(stats); // Prints out an instance of a node.js fs.Stats object
    }
  });
}

readXlsxFile(path.join(__dirname, './bak/832.xlsx'), { sheet: origin }).then(row => {
  const len = row.length
  let result = []
  for (let i = 1; i < len; i++) {
    if (row[i][7]) {
      if (!row[i][11]) {
        if (/元/.test(row[i][12])) {
          let tem = row[i][12].replace(/[^0-9]/ig, '')
          tem = Number(tem)
          let cost = Math.round(row[i][9] || row[i][10])
          if (cost > 0) {
            result.push({
              name: `${row[i][5]}-${row[i][7]}`,
              cost: Math.round(row[i][9] || row[i][10]),
              sell: Math.round(tem),
              origin: origin
            })
          }
        } else {
          let cost = Math.round(row[i][9] || row[i][10])
          if (cost > 0) {
            result.push({
              name: `${row[i][5]}-${row[i][7]}`,
              cost: Math.round(row[i][9] || row[i][10]),
              sell: isNaN(row[i][12]) ? 0 : Math.round(row[i][12]),
              origin: origin
            })
          }
        }
      }
    }
  }
  for (let i = 0; i < result.length; i++) {
    if (/null/gi.test(result[i].name)) {
      let tem = result[i].name.split('-')
      result[i].name = `${result[i - 1].name.split('-')[0]}-${tem[1]}`
    }
  }

  genXlsx(ws, result)
})
