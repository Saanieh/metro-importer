const XLSX = require('xlsx')
const chalk = require('chalk')
const program = require('commander')


program
    .version('0.1.0')
    .option('-e, --to-end', 'Is to end')
    .option('-d, --work-day [day]', 'Work Day')
    .option('-p, --path [path]', 'File path')
    .parse(process.argv);


console.log(chalk.green('Iran Metro - Excel Parser'))

let isToEnd = 0
if (program.toEnd) {
    isToEnd = 1
}

console.log(chalk.yellow(`is to end set to ${isToEnd}`))

let workDay = 0
if (program.workDay) {
    workDay = parseInt(program.workDay)
}

switch (workDay) {
    case 0:
        console.log(chalk.yellow(`This excel is set for Saturday to Wednesday.`))
        break
    case 1:
        console.log(chalk.yellow(`This excel is set for Thursday.`))
        break
    case 2:
        console.log(chalk.yellow(`This excel is set for Friday.`))
        break
}


if (program.path) {
    console.log(chalk.blue(`File path passed is ${program.path}`))
    let workbook = XLSX.readFile(program.path);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]]
    let data = XLSX.utils.sheet_to_json(worksheet)
    let stationTimes = []

    data.splice(0, 1)
    for (let record of data) {
        let stations = Object.keys(record)
        for (let station of stations) {
            stationTimes.push({ station_pid: station, time: record[station], is_to_end: isToEnd, workday: workDay, description: null, is_fast: 0 })
        }
    }

    let newSheet = XLSX.utils.json_to_sheet(stationTimes)
    let newExcel = { SheetNames: ['newFile'], Sheets: { newFile: newSheet}}
    XLSX.writeFile(newExcel, 'newFile.csv')
}

