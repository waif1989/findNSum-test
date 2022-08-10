const path = require('path')
const fs = require('fs');
const xl = require('excel4node');
const readXlsxFile = require('read-excel-file/node')

const SELL_PLAN = 300 // 套餐价格
const TARGET_PROFIT = 0.26 // 利润率
const PRODUNTS_NUMS = 5 // 套餐需要多少种商品
const BIAS = 10 // 利润偏差值
const MAX_LIST_LIMIT = 1000 // 循环最大条数，多出会导致内存溢出
const XLSX_LIST = 100000 // 方案输出
let targetCost = SELL_PLAN - (SELL_PLAN * TARGET_PROFIT) // 达到套餐价格指定利润率的目标成本价格

const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');
ws.cell(1, 1).string('商品编码')
ws.cell(1, 2).string('商品名称')
ws.cell(1, 3).string('商品成本')
ws.cell(1, 4).string('市场价')
ws.cell(1, 5).string('产地')

// const writer = fs.createWriteStream(path.join(__dirname, './allProducts.js'));
// writer.write(JSON.stringify(allProducts));

function swap(A, i, j) {
  const t = A[i];
  A[i] = A[j];
  A[j] = t;
}

/**
 *
 * @param {*} A  数组
 * @param {*} p  起始下标
 * @param {*} r  结束下标 + 1
 */
function divide(A, p, r) {
  const x = A[r - 1].cost;
  let i = p - 1;

  for (let j = p; j < r - 1; j++) {
    if (A[j].cost <= x) {
      i++;
      swap(A, i, j);
    }
  }

  swap(A, i + 1, r - 1);

  return i + 1;
}

/**
 *
 * @param {*} A  数组
 * @param {*} p  起始下标
 * @param {*} r  结束下标 + 1
 */
function qsort(A, p = 0, r) {
  r = r || A.length;

  if (p < r - 1) {
    const q = divide(A, p, r);
    qsort(A, p, q);
    qsort(A, q + 1, r);
  }

  return A;
}

/**
 * @param {number[]} nums
 * @param {number} target
 * @return {number[][]}
 */
const NSum = function (nums, target, TargetN, cb) {
  const findNSum = function (nums, target, N, result, results) {
    if (nums.length < N || target < nums[0].cost * N || target > nums[nums.length - 1].cost * N) {
      return;
    }
    if (N === 2) {
      let l = 0,
        r = nums.length - 1;
      while (l < r) {
        let s = nums[l].cost + nums[r].cost;
        if (s >= (target - BIAS) && s <= target) {
          results.push(result.concat([nums[l], nums[r]]));
          while (l < r && nums[l].cost === nums[l + 1].cost) {
            l++;
          }
          while (r > l && nums[r].cost === nums[r - 1].cost) {
            r++;
          }
          l++;
          r--;
        } else if (s < (target - BIAS)) {
          l++;
        } else {
          r--;
        }
      }
    } else {
      for (let i = 0; i < nums.length - N + 1; i++) {
        if (i === 0 || (i > 0 && nums[i - 1].cost !== nums[i].cost)) {
          findNSum(nums.slice(i + 1), target - nums[i].cost, N - 1, result.concat([nums[i]]), results);
        }

      }
    }
  };
  let results = [];
  /*nums.sort(function (a, b) {
    return a - b;
  });*/
  findNSum(nums, target, TargetN, [], results);
  return results;
};

const genXlsx = (ws, arr) => {
  const len = arr.length > XLSX_LIST ? XLSX_LIST : arr.length
  for (let i = 0; i < len; i++) {
    for (let j = 0; j < arr[i].length; j++) {
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 1).string(String(arr[i][j].id))
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 2).string(String(arr[i][j].name))
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 3).number(Number(arr[i][j].cost))
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 4).string(String(arr[i][j].sell))
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 5).string(String(arr[i][j].origin || ''))
    }
  }

  wb.write(path.join(__dirname, './ExcelFile.xlsx'), function(err, stats) {
    if (err) {
      console.error(err);
    } else {
      console.log(stats); // Prints out an instance of a node.js fs.Stats object
    }
  });
}

const run = (removeDuplicates = true) => {
  readXlsxFile(path.join(__dirname, './bak/test3.xlsx'), { sheet: 'Sheet1' }).then((rows) => {
  // readXlsxFile(path.join(__dirname, './test2.xlsx'), { sheet: '标准化方案汇总' }).then((rows) => {
    const productIds = []
    const packagesIds = []
    const packages = []
    let allProducts = [] // 所有不重复的商品
    let avgPackagesCost = 0
    let packagesCost = 0
    const rowLen = rows.length > MAX_LIST_LIMIT ? MAX_LIST_LIMIT : rows.length
    for (let i = 1; i < rowLen; i++) {
      if (removeDuplicates) {
        if (/纸箱/gi.test(rows[i][1]) || /物料/gi.test(rows[i][0]) || /包材预留/gi.test(rows[i][1]) || /环保袋/gi.test(rows[i][1])) {
          if (!packagesIds.includes(rows[i][1]) && !packagesIds.includes(rows[i][0])) {
            packagesIds.push(rows[i][1])
            packages.push({
              id: rows[i][0],
              name: rows[i][1],
              cost: rows[i][2],
            })
            packagesCost = packagesCost + rows[i][2]
          }
        } else if (!productIds.includes(rows[i][1])) {
          productIds.push(rows[i][1])
          allProducts.push({
            id: rows[i][0],
            name: rows[i][1],
            cost: Math.round(rows[i][2] || 0),
            sell: rows[i][3],
          })
        }
      } else {
        let cost = Math.round(rows[i][2] || 0)
        if (cost <= SELL_PLAN && cost >= 25) {
          allProducts.push({
            id: rows[i][0],
            name: rows[i][1],
            cost: Math.round(rows[i][2] || 0),
            sell: rows[i][3],
            origin: rows[i][5],
          })
        }
      }
    }

    avgPackagesCost = packagesCost / packages.length // 包装平均成本
    allProducts = qsort(allProducts)
    const allAssemble = NSum(allProducts, targetCost, PRODUNTS_NUMS)
    // console.log('-----', allAssemble)

    genXlsx(ws, allAssemble)


  })
}

run(false)
