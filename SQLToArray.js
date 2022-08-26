const arrayToXlsx = async (array, namefile) => { //array: array de objetos, namefile: nombre del archivo
    const workbook = require("xlsx").utils.book_new();// Create a new workbook
    require("xlsx").utils.book_append_sheet(workbook, // Add the worksheet to the workbook with the name "HojaDeCalculoDelLibro"
        require("xlsx").utils.json_to_sheet(array), // generate worksheet from array
        "Hoja1");  // Add the worksheet to the workbook with the name "HojaDeCalculoDelLibro"
    require("xlsx").writeFile(workbook, namefile);/* create an XLSX file and try to save to namefile.xlsx */
};
exports.arrayToXlsx = arrayToXlsx;