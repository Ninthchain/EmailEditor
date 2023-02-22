const ExcelJs = require('exceljs');
const clc = require('cli-color')
const { abort } = require('process');
const FS = require('fs')

const banDomains = [
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
function isCorrectEmail(email)  {
    if(email == null) return;
    for(let slotIt = 0 ; slotIt < banDomains.length; slotIt++) {
        
        if(email.includes(banDomains[slotIt])) return false;
    }
    return true;
}



function logError(str) {
    console.log(clc.red(str))
    abort()
}
function logComment(str) { 
    console.log(clc.cyan(str))
}
function logWarning(str) {
    console.log(clc.yellowBright(str))
}
function logSuccess(str) {
    console.log(clc.green(str))
}

logComment('Copying original file..\n')
try {
    logComment('Copying started')
    FS.copyFileSync('./input.xlsx', 'inputCopy.xlsx')
}
catch(err) {
    logError('Error while copying')
} 
logSuccess('Copy done!')

logComment('Editing started')

let workbook = new ExcelJs.Workbook();

workbook.xlsx.readFile('./input.xlsx').then(async function(wb) {
    
    const workAdressesColumn = 6;
    wb.eachSheet((sheet) => {

      sheet.eachRow({includeEmpty: false}, async (row) => {
        
        if(!isCorrectEmail(row.getCell(workAdressesColumn).value)) {

            var changed = false;
            
            logComment(`Found email with banned domain.\n\ Address:${row.getCell(workAdressesColumn).value} Excel position: ${row.getCell(workAdressesColumn).$col$row}`)
            logComment('Trying to find additional address')
            
            for(let slotIt = 7; slotIt != 15; slotIt++) {

                if(row.getCell(slotIt).value == null) return;

                if(row.getCell(slotIt).value.length > 0) {
                    
                    if(!isCorrectEmail(row.getCell(slotIt).value)) continue;
                        try {
                            row.getCell(workAdressesColumn).value = row.getCell(slotIt).value
                        }
                        catch(err) {
                            logError(err)
                        }}
                        
                        logSuccess(`The email has successfully changed to ${row.getCell(slotIt).value}`)
                        changed = true
                        break;
                }
            } if(!changed) logWarning(`The email ${row.getCell(workAdressesColumn).value} have not been replaced. Dont have any right addresses`)
        })
      })
    return workbook.xlsx.writeFile('output.xlsx')
})
