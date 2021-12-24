(async () => {
  const xlsx = require('node-xlsx')
  const moment = require('moment')
  const fs = require('fs')
  const { title } = require('process')
  
  // 获取开班日期
  // const START_DATE = process.env.START_DATE 
  process.stdout.write('请输入起始日期: 格式例如 20210102 或 2021-12-24:\n\n')
  process.stdout.write('起始日期为： ')
  
  process.stdin.setEncoding('utf8')
  const START_DATE = await readlineSync()

  // 日期检测
  if (!START_DATE) {
    console.log('您忘了输入起始日期了 \n')
    process.exit()
  }

  const len = START_DATE.length
  if (
    !START_DATE ||
    (len !== 8 && len !== 10) ||
    (len === 8 && /\D/.test(START_DATE)) ||
    (len === 10 && /[^0-9-]/.test(START_DATE)) ||
    !moment(START_DATE).isValid()
  ) {
    console.log('连个日期都输不对，还能不能行了，看完 README.md 再来吧（嫌弃脸）\n')
    process.exit()
  }
  
  // 读取文件
  const sheets = xlsx.parse('./temp.xlsx', { cellDates: true })
  
  // 处理数据
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
  fs.writeFileSync(`./dist/新生成的学习计划${ moment().format('YYYY-MM-DD HH时mm分ss秒') }.xlsx`, buffer, { flag: 'w'})
  
  process.on('exit', () => {
    console.log('======================================= \n\n学习计划生成成功，请到dist目录下查看...\n')
  });
  
  function formatNR (str) {
    return str && str.replace(/\r\n/g, '\n')
  }
  
  function readlineSync() {
    return new Promise((resolve, reject) => {
      process.stdin.resume()
      process.stdin.on('data', function (data) {
        process.stdin.pause()
        resolve(data.trim())
      })
    })
  }
})()