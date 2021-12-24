const xlsx = require('node-xlsx')
const moment = require('moment')
const fs = require('fs')

// 读取文件
const sheets = xlsx.parse('./temp.xlsx', { cellDates: true })
// 获取开班日期
const START_DATE = process.env.START_DATE

const Part1 = [['日期', '学习目标', '对应视频']]

sheets.forEach(sheet => {
  const data = sheet.data
  const len = data.length
  for(let i = 1; i < len; i++) {
    const row = data[i]
    if(row && row.length){
      Part1.push([
        moment(START_DATE).add(i, 'day').format('YYYY/MM/DD'), // 日期
        formatNR(row[1]), // 学习目标
        formatNR(row[2]) // 对应视频
      ])
    }
  }
})

// 写入文件
let buffer = xlsx.build([{
  name: 'Part1',
  data: Part1
}], {
  sheetOptions: {
    '!cols': [{ wch: 20 }, { wch: 45 }, { wch: 35 }]
  }
})
fs.writeFileSync('./新生成的学习计划.xlsx', buffer, { flag: 'w'})

function formatNR (str) {
  return str && str.replace(/\r\n/g, '\n')
}
