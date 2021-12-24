# 说明
该工具用于根据输入日期自动生成每日学习目标，运行程序需要 [node.js](https://nodejs.org/zh-cn/) 环境支持。

使用方式：
1. 双击 .bat 运行后，输入第一天的日期，程序会自动在 dist 目录下生成后续的学习计划，文件名为 “新生成的学习计划.xlsx”。
2. 日期格式
  - 日期为个位数时，十位需要用0表示，如：
    - 2022 年 1 月 6 日表示为 20220106
  - 年月日还可以用 - 连接，如：
    - 2021 年 12 月 24 日表示为 2021-12-24