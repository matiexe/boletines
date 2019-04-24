var docactual = SpreadsheetApp.getActiveSpreadsheet();
var sheet= docactual.getActiveSheet();
var cell =docactual.getRangeByName('B2');
var rule;
var rows = sheet.getDataRange();
var value =rows.getValues();

//Esta Funcion crea un listado desplegable que muestra los alumnos segun el curso en que esten
function Validacion() {
  var nhojas =docactual.getSheets();//obtengo todas las hojas del libro
  var cursoactual1= nhojas[0].getName();//aacedo a la primera hoja del libro y obtiene el nombre
  var cursoactual2= nhojas[1].getName();
  var ncurso = value [2][1]; //se accede al valor del curso de la planilla del alumno
  if(ncurso==cursoactual2){// se evalua si el valor de la celda es igual al curso del alumno
    var curso = docactual.getSheetByName(cursoactual2);
    var rango = curso.getRange('B7:B30');
    rule = SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).build();//se crea la regla con el rango provisto
    cell.setDataValidation(rule);
     
     }
  else{
    if (ncurso ==cursoactual1){
      var curso = docactual.getSheetByName(cursoactual1);
      var rango = curso.getRange('B7:B30');
      rule = SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).build();
      cell.setDataValidation(rule);
      
    }
  
  }
   

}

