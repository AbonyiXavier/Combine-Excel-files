let xlsx = require("xlsx")
let fs = require("fs")
let path = require("path")

let sourceDir = "Files"  // add your excel files on the Files folder and run the script
function readFileToJson (filename){
    
    let wb = xlsx.readFile(filename,{cellDates:true})
    
    let firstTabName = wb.Props.SheetNames[0] // Get the first tab name of the excel sheet
    
    let ws = wb.Sheets[firstTabName]
    
    //Take the worksheet and read the data to an array
    let data = xlsx.utils.sheet_to_json(ws)  //  convert the workshet object to an array
    
   return data

}
// let read = readFileToJson("no-transaction-1.xlsx")
let targetDir = path.join(__dirname, sourceDir); //Getting the files from the folder path
let files = fs.readdirSync(targetDir) //  reading the excel files from the folder path
// console.log(files);

let combinedData = [];

files.forEach(function(file) {
    let fileExtension = path.parse(file).ext // Get file extension to be .xlsx
    if (fileExtension === ".xlsx" && file[0] !== "~") { // valiadating to make sure all file gotten are excel and no file extension on the front e.g ~
        let fullFilePath = path.join(__dirname, sourceDir, file);
        // console.log(fullFilePath);
        let readData = readFileToJson(fullFilePath)
        combinedData = combinedData.concat(readData)
    }

})

// console.log(combinedData.length); get the total length of the combined data

let newWB = xlsx.utils.book_new(); // new work book
let newWS = xlsx.utils.json_to_sheet(combinedData) // new worksheet

xlsx.utils.book_append_sheet(newWB, newWS, "combined data")

xlsx.writeFile(newWB, "mikrocombineddata.xlsx")

console.log("done!");
