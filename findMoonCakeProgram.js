const path = require('path')
const fs = require('fs');
const xl = require('excel4node');
const readXlsxFile = require('read-excel-file/node')

const SHEET_NAME = '标准化方案汇总'
const SELL_PLAN = 300 // 套餐价格
const TARGET_PROFIT = 0.26 // 利润率
const PRODUNTS_NUMS = 5 // 套餐需要多少种商品
const BIAS = 10 // 利润偏差值
const MAX_LIST_LIMIT = 1000 // 商品列表上一次性最多计算的数量，多出这个数会导致内存溢出
const XLSX_LIST = 200 // 方案输出
let targetCost = SELL_PLAN - (SELL_PLAN * TARGET_PROFIT) // 达到套餐价格指定利润率的目标成本价格

const wb = new xl.Workbook(); // 生成xlsx表格对象
const ws = wb.addWorksheet('Sheet 1'); // 默认生成一张表格，名字为Sheet 1
ws.cell(1, 1).string('商品编码') // 第一行第一列
ws.cell(1, 2).string('商品名称') // 第一行第二列
ws.cell(1, 3).string('商品成本') // 第一行第三列
ws.cell(1, 4).string('市场价') // 第一行第四列
ws.cell(1, 5).string('产地') // 第一行第五列

/**
 * 快速排序
 * @param {*} A  数组
 * @param {*} p  起始下标
 * @param {*} r  结束下标 + 1
 */
const qSortFn = (A, p = 0, r) => {
  const swap = (A, i, j) => {
    const t = A[i];
    A[i] = A[j];
    A[j] = t;
  }
  
  const divide = (A, p, r) => {
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
  
  const qsort = (A, p = 0, r) => {
    r = r || A.length;
    
    if (p < r - 1) {
      const q = divide(A, p, r);
      qsort(A, p, q);
      qsort(A, q + 1, r);
    }
    
    return A;
  }
  
  return qsort(A, p = 0, r)
}

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
  findNSum(nums, target, TargetN, [], results);
  return results;
};

/**
 * 生成结果表格的方法
 * */
const genXlsx = (ws, arr) => {
  const len = arr.length > XLSX_LIST ? XLSX_LIST : arr.length
  for (let i = 0; i < len; i++) {
    for (let j = 0; j < arr[i].length; j++) {
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 1).string(String(arr[i][j].id)) // 商品编码
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 2).string(String(arr[i][j].name)) // 商品名称
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 3).number(Number(arr[i][j].cost)) // 商品成本
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 4).string(String(arr[i][j].sell)) // 市场价
      ws.cell((i * (PRODUNTS_NUMS + 1) + 1) + (j + 1), 5).string(String(arr[i][j].origin || '')) // 产地
    }
  }
  
  wb.write(path.join(__dirname, './output/result.xlsx'), function(err, stats) {
    if (err) {
      console.error(err);
    } else {
      console.log(stats); // Prints out an instance of a node.js fs.Stats object
    }
  });
}

/**
 * 程序主方法
 * */
const run = () => {
  readXlsxFile(path.join(__dirname, './input/20220805.xlsx'), { sheet: SHEET_NAME }).then(rows => {
    const len = rows.length > MAX_LIST_LIMIT ? MAX_LIST_LIMIT : rows.length
    for (let i = 0; i < len; i++) {
    
    }
  })
}