//Agregar librerias realizadas
const mssql = require('mssql/msnodesqlv8')
const fs = require('fs').promises;//Importar codigo de manejo de filestream
const toXLSX = require('./SQLToArray').arrayToXlsx;//Importar codigo de exportación de excel
const { exec } = require("child_process");//cmd
//constantes
const pathin = './sql/'; //Carpeta de archivos sql para exportar a xlsxworker
const pathout = './exportXLS/';//Carpeta de archivos sql para exportar a xlsxworker
const server = 'KUP\\SQLEXPRESS' //Servidor e instancia de base de datos
const database = 'master'//Nombre de base de datos
//Crear agrupamiento de conexiones: https://es.wikipedia.org/wiki/Connection_pool
const db = new mssql.ConnectionPool({ database: database, server: server, options: { trustedConnection: true } })//Autenticación de windows a la base de datos
async function execute() { //Función asincronica
    try { await fs.mkdir(pathout); console.log("Creada la carpeta " + pathout) } catch { console.log('Carpeta Existente ' + pathout) } //Intente Crear Carpeta de archivos xlsx y Si Existe no haga nada
    try { await fs.mkdir(pathin); console.log("Creada la carpeta " + pathin) } catch { console.log('Carpeta Existente ' + pathin) } //Intente Crear Carpeta de archivos sql y Si Existe no haga nada
    try {
        await db.connect();//Espere a que conecte a sql para continuar
        let sqls = [];//Array de promesas de resultados de sql
        let files = [];//Nombres De Archivos
        for (let file of (await fs.readdir(pathin))) {//Para Cada Archivo del directorio haga
            files.push(file);//Almacena el nombre de cada archivo
            sqls.push(db.request().query(await fs.readFile(pathin + file, 'utf8')));//Ejecuta consulta sql del archivo y continua con el demás código sin esperar la respuesta inmediatamente haciendo que varias consutlas se ejecuten en paralelo
        }
        if ((await fs.readdir(pathin)).length == 0) {
            console.log("No existen archivos a exportar por favor agregue archivos con consultas sql en la carpeta: " + pathin.replace("/", "").replace("/", "").replace(".", ""))
            console.log("Abrire la carpeta donde debe poner consultas sql por usted");            
            exec("explorer " + (__dirname + "\\" + pathin.replace("/", "").replace("/", "").replace(".", "")));
            for (let file of (await fs.readdir(pathout)))
                await fs.unlinkSync(file)
            console.log("Finalizado puede cerrar la ventana")
        } else {
            console.log("Archivos sql a convertir sus resutados en xlsx");
            console.dir((await fs.readdir(pathin)))
        }
        for (let [i, r] of (await Promise.all(sqls)).entries()) //Para las respuestas de todas las consultas sql exporte a archivos xlsx con el nombre del archivo sql
            toXLSX(r.recordset, pathout + files[i].substring(0, files[i].length - 4) + '.xlsx');
        console.log('Archivos Exportados Exitosamente')
        console.dir((await fs.readdir(pathout)))
        console.log("Finalizado puede cerrar la ventana")
        console.log("Abrire la carpeta por usted en donde se exportaron los archivos de hoja(s) de calculo");
        exec("explorer " + (__dirname + "\\" + pathout.replace("/", "").replace("/", "").replace(".", "")));
        db.close();//cierre la unica conexión que se abrio para todas las consultas 
    } catch { }
}
execute();//Inicia la ejecución de código
