const ExcelJs = require('exceljs');
const clc = require('cli-color')
const { abort, exit } = require('process');
const FS = require('fs')
const { isStringObject } = require('util/types');
const { notEqual } = require('assert');
const { isNull } = require('util');

const banWords = [
    'noreply', 
    'mail', 
    'post', 
    'office', 
    'support', 
    'opt', 
    'tender', 
    'info', 
    'admin', 
    'opt'
]


/**
 * @param {String} email
 * @return {Boolean}
 */
function isIncorrectEmail(email)  {
    if(email == null) return;
    for(let i = 0 ; i < banWords.length; i++) {
        
        if(email.includes(banWords[i])) return true;
    }
    return false;
}
console.log(clc.cyan('Делаем копию excel файла...'))
try {
    FS.copyFileSync('./input.xlsx', 'inputCopy.xlsx')
}
catch(err) {
    console.error(clc.red('Ошибка в копировании'));
    abort()
} 
console.log(clc.green('Копия сделана\n'))

// email columns from 4th index to 15th number;
let workbook = new ExcelJs.Workbook();
  workbook.xlsx.readFile('./input.xlsx').then(async function(wb){
    const workAdressesColumn = 6;
    wb.eachSheet((sheet, id) => {

      sheet.eachRow({includeEmpty: false}, async (row, number) => {
        if(isIncorrectEmail(row.getCell(workAdressesColumn).value)) {
            console.log(clc.cyanBright(`Обнаружен адрес эл почты с запрещеным доменом.\n\tАдрес:${row.getCell(workAdressesColumn).value} Позиция ячейки: ${row.getCell(workAdressesColumn).$col$row}`))
            console.log(clc.green(`Ищу доступный дополнительный адрес`))
            var previousValue = row.getCell(workAdressesColumn).value
            var changed = false;
            for(let i = 7; i != 15; i++) {
                if(row.getCell(i).value == null) return;
                if(row.getCell(i).value.length > 0) {
                    if(isIncorrectEmail(row.getCell(i).value)) continue;
                        try {
                            row.getCell(workAdressesColumn).value = row.getCell(i).value
                        }
                        catch(err) {
                            console.error(clc.red(err));
                            abort()
                        }}
                        console.log(clc.green('Адрес ') + clc.whiteBright(previousValue) + clc.green(' заменён успешно на ') + clc.whiteBright(row.getCell(i).value))
                        changed = true
                        break;
                }
            }
            if(!changed) console.warn(clc.red(`Адрес эл почты ${row.getCell(workAdressesColumn).value} неудалось заменить. Нету дополнительных адресов для замены`))
        })
      })
      return workbook.xlsx.writeFile('output.xlsx')
    })
