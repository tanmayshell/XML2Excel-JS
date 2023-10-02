import { convertXmlToExcel } from './xmlToExcel.js';
const xmlFilePath='task.xml';
const excelFilePath= 'task.xlsx'

convertXmlToExcel(xmlFilePath,excelFilePath,(err, result)=>{
    if(err){
        console.error('error', err)
    } else {
        console.log('result', result)
    }
})