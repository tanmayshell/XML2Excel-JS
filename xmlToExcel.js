import fs from 'fs'
import xml2js from 'xml2js'
import xlsx from 'xlsx'

export function convertXmlToExcel(xmlFilePath, excelFilePath, cb){
    fs.readFile(xmlFilePath, 'utf-8', (err, xmlData) => {
        if(err){
            cb(err)
            return
        }
        const parser= new xml2js.Parser({explicitArray: false, mergeAttrs:true});
        parser.parseString(xmlData, (err, jsObject)=>{
            if(err){
                cb(err)
                return
            }

            try {
                const moviesData=jsObject.binge.movies.movie
                const seriesData=jsObject.binge.series.serie
                const moviesWs = xlsx.utils.json_to_sheet(moviesData);
                const seriesWs=xlsx.utils.json_to_sheet(seriesData);              
                const wb = xlsx.utils.book_new();

                function formatRange(sheet, range, style){
                    for(let row=range.s.r; row<=range.e.r; row++){
                        for(let col=range.s.c; col<=range.e.c; col++){
                            const reference=xlsx.utils.encode_cell({r:row, c:col});
                            if(!sheet[reference]) continue;
                            sheet[reference].s= style
                        }

                    }
                }



                xlsx.utils.book_append_sheet(wb, moviesWs, "Movies");
                xlsx.utils.book_append_sheet(wb, seriesWs, "Series");

                const bold= {font:{bold:true}}
                const headerRange=xlsx.utils.decode_range(moviesWs['!ref'])
                formatRange(moviesWs, headerRange, bold)

                xlsx.writeFile(wb, excelFilePath)
                cb(null, 'Success')
            }
            catch(error){
                cb(error)
            }
        })
    })
}