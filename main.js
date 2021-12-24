const xlsx = require('node-xlsx')
const moment = require('moment')
const fs = require('fs')
const { title } = require('process')

// 读取文件
const sheets = xlsx.parse('./temp.xlsx', { cellDates: true })
// 获取开班日期
const START_DATE = process.env.START_DATE

const titles = ['日期', '学习目标', '对应视频']
const sheetsLists = []
let j = 0
sheets.forEach(sheet => {
  const sheetData = [titles.slice()]
  const data = sheet.data
  const len = data.length
  let i
  for(i = 1; i < len; i++,j++) {
    const row = data[i]
    if(row && row.length){
      sheetData.push([
        moment(START_DATE).add(j, 'day').format('YYYY/MM/DD'), // 日期
        formatNR(row[1]), // 学习目标
        formatNR(row[2]) // 对应视频
      ])
    }
  }
  sheetsLists.push({
    name: sheet.name,
    data: sheetData
  })
})

// 写入文件
let buffer = xlsx.build(sheetsLists, {
  sheetOptions: { '!cols': [{ wch: 20 }, { wch: 45 }, { wch: 35 }] }
})
fs.writeFileSync('./新生成的学习计划.xlsx', buffer, { flag: 'w'})

function formatNR (str) {
  return str && str.replace(/\r\n/g, '\n')
}
