  /**
    \maingpage Boletin Austro
    \author Urra Gauna Matias
    \date Abril-2018
  */
  var boletin = SpreadsheetApp.getActive();
  var contador= boletin.getActiveSheet().getRange("A48").getValue();
  var nom_full_estat= "estat";

/**
  \fn nombrePestania() Cambia el nombre de una pestaña en la hoja de calculo
  \var hoja obtiene la hoja activa de libro
       rows obtiene las c 
*/
function nombrePestania() {
  var hoja = boletin.getActiveSheet();
  var rows =hoja.getDataRange();
  var value = rows.getValues();
  var nombre1 = value[1][2];
  var nombre2 = value[1][26];
  hoja.setName(nombre1+" "+nombre2);
  var alumno = value[1][1];
  hoja.setName(alumno);
  SpreadsheetApp.flush();
  
}

function onOpen() {
  SpreadsheetApp.getUi()
       .createMenu('Boletin Austro')
       .addItem('GENERAR BOLETIN', 'generarBoletines')
       .addSeparator()
       .addItem('Eliminar Alumno', 'eliminaHoja')
       /*.addSubMenu(SpreadsheetApp.getUi().createMenu('My sub-menu')
           .addItem('One sub-menu item', 'mySecondFunction')
           .addItem('Another sub-menu item', 'myThirdFunction'))*/
         .addItem('Promedios','buscaPromedios')
         .addItem('Proteger','showDialog')
         .addItem('Eliminar Proteccion','eliminar')
       .addToUi();
}
function contadorAlumnos(contador){ 
  var nhojas= boletin.getSheets();
  var hojactiva=boletin.getActiveSheet();
  var total = boletin.getNumSheets();
  var cursoactual1=nhojas[2].getName();//obtengo el nombre de la hoja del primer curso
  var cursoactual2=nhojas[3].getName();//obtengo el nombre de la hoja de segundo curso
  var curso=hojactiva.getRange("B3").getValue();  //obtengo el valor del curso en la hoja del alumno
  var posicion = hojactiva.getRange('A48').getValue();
  if(curso ==cursoactual1){
    var  cantalumn = nhojas[1].getRange("A50").getValue();
    if(contador<cantalumn){
      posicion=posicion+contador;
      hojactiva.getRange("A48").setValue(posicion);
      }
    else{
      posicion =7;
      hojactiva.getRange("A48").setValue(posicion);
    }
  }
  else{
  
    if(curso ==cursoactual2){
    var  cantalumn = nhojas[2].getRange("A50").getValue();
    if(contador<cantalumn){
      posicion=posicion+contador;
      hojactiva.getRange("A48").setValue(posicion);
      }
    else{
      posicion=7;
      hojactiva.getRange("A48").setValue(posicion);
      }
    
  }
  }
}
function eliminaHoja(){
  boletin.deleteActiveSheet();
  var nhojas = boletin.getNumSheets();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet()
  var hojactiva = boletin.getActiveSheet();
  contador=contador-1;
  hojactiva.getRange("A48").setValue(contador);
  hojactiva.getRange("A49").setValue(nhojas);
  

}
function setMaterias(){
  var activeSheet = boletin.getActiveSheet();
  var orientacionUno = boletin.getSheets()[2].getName();
  var orientacionDos = boletin.getSheets()[3].getName();
  var orientacionActual = activeSheet.getRange('B3').getValue();
  var columaMaterias=["C","J","Q","X","AE","AL","AS","AZ","BG","BN","BU","CB","CI","CP"]
  var posicion = 6
  if (orientacionActual == orientacionUno){
      for(var i =0 ;i<14;i++){
        posicion=posicion+1
        activeSheet.getRange("B"+posicion).setFormula("='"+orientacionActual+"'!"+columaMaterias[i]+"4");
        activeSheet.getRange("A"+posicion).setFormula("='"+orientacionActual+"'!"+columaMaterias[i]+"5");
      }
    }
   else{
     for(var i =0 ;i<14;i++){
        posicion=posicion+1
        activeSheet.getRange("B"+posicion).setFormula("='"+orientacionDos+"'!"+columaMaterias[i]+"4");
        activeSheet.getRange("A"+posicion).setFormula("='"+orientacionDos+"'!"+columaMaterias[i]+"5");
      }
   }
  
}
  //Setear Materias segun orientacion
function setNotas(){
  var hojactiva=boletin.getActiveSheet();
  var nhojas=boletin.getSheets();
  var cursoactual1=nhojas[2].getName();
  var cursoactual2=nhojas[3].getName();
  var curso=hojactiva.getRange("B3").getValue();
  var contador = hojactiva.getRange('A48').getValue();
  if(curso ==cursoactual1)
  {
    //MATEMATICA
    hojactiva.getRange("C7").setFormula("='"+cursoactual1+"'!C$"+contador);//Primer Trimestre
    hojactiva.getRange("E7").setFormula("='"+cursoactual1+"'!D$"+contador);//Segundo Trimestre
    hojactiva.getRange("G7").setFormula("='"+cursoactual1+"'!E$"+contador);//Tercer Trimestre
    hojactiva.getRange("I7").setFormula("='"+cursoactual1+"'!F$"+contador);//Diciembre
    hojactiva.getRange("K7").setFormula("='"+cursoactual1+"'!G$"+contador);//Febrero
    hojactiva.getRange("M7").setFormula("='"+cursoactual1+"'!H$"+contador);//Mesa
    hojactiva.getRange("O7").setFormula("='"+cursoactual1+"'!I$"+contador);//Nota Final
    //lENGUA
    hojactiva.getRange("C8").setFormula("='"+cursoactual1+"'!J$"+contador);
    hojactiva.getRange("E8").setFormula("='"+cursoactual1+"'!K$"+contador);
    hojactiva.getRange("G8").setFormula("='"+cursoactual1+"'!L$"+contador);
    hojactiva.getRange("I8").setFormula("='"+cursoactual1+"'!M$"+contador);
    hojactiva.getRange("K8").setFormula("='"+cursoactual1+"'!N$"+contador);
    hojactiva.getRange("M8").setFormula("='"+cursoactual1+"'!O$"+contador);
    hojactiva.getRange("O8").setFormula("='"+cursoactual1+"'!P$"+contador);
     //INGLES
    hojactiva.getRange("C9").setFormula("='"+cursoactual1+"'!Q$"+contador);
    hojactiva.getRange("E9").setFormula("='"+cursoactual1+"'!R$"+contador);
    hojactiva.getRange("G9").setFormula("='"+cursoactual1+"'!S$"+contador);
    hojactiva.getRange("I9").setFormula("='"+cursoactual1+"'!T$"+contador);
    hojactiva.getRange("K9").setFormula("='"+cursoactual1+"'!U$"+contador);
    hojactiva.getRange("M9").setFormula("='"+cursoactual1+"'!V$"+contador);
    hojactiva.getRange("O9").setFormula("='"+cursoactual1+"'!W$"+contador);
     //EDUCACION FISICA
    hojactiva.getRange("C10").setFormula("='"+cursoactual1+"'!X$"+contador);
    hojactiva.getRange("E10").setFormula("='"+cursoactual1+"'!Y$"+contador);
    hojactiva.getRange("G10").setFormula("='"+cursoactual1+"'!Z$"+contador);
    hojactiva.getRange("I10").setFormula("='"+cursoactual1+"'!AA$"+contador);
    hojactiva.getRange("K10").setFormula("='"+cursoactual1+"'!AB$"+contador);
    hojactiva.getRange("M10").setFormula("='"+cursoactual1+"'!AC$"+contador);
    hojactiva.getRange("O10").setFormula("='"+cursoactual1+"'!AD$"+contador);
     //FEC
    hojactiva.getRange("C11").setFormula("='"+cursoactual1+"'!AE$"+contador);
    hojactiva.getRange("E11").setFormula("='"+cursoactual1+"'!AF$"+contador);
    hojactiva.getRange("G11").setFormula("='"+cursoactual1+"'!AG$"+contador);
    hojactiva.getRange("I11").setFormula("='"+cursoactual1+"'!AH$"+contador);
    hojactiva.getRange("K11").setFormula("='"+cursoactual1+"'!AI$"+contador);
    hojactiva.getRange("M11").setFormula("='"+cursoactual1+"'!AJ$"+contador);
    hojactiva.getRange("O11").setFormula("='"+cursoactual1+"'!AK$"+contador);
     //HISTORIA
    hojactiva.getRange("C12").setFormula("='"+cursoactual1+"'!AL$"+contador);
    hojactiva.getRange("E12").setFormula("='"+cursoactual1+"'!AM$"+contador);
    hojactiva.getRange("G12").setFormula("='"+cursoactual1+"'!AN$"+contador);
    hojactiva.getRange("I12").setFormula("='"+cursoactual1+"'!AO$"+contador);
    hojactiva.getRange("K12").setFormula("='"+cursoactual1+"'!AP$"+contador);
    hojactiva.getRange("M12").setFormula("='"+cursoactual1+"'!AQ$"+contador);
    hojactiva.getRange("O12").setFormula("='"+cursoactual1+"'!AR$"+contador);
     //GEOGRAFIA
    hojactiva.getRange("C13").setFormula("='"+cursoactual1+"'!AS$"+contador);
    hojactiva.getRange("E13").setFormula("='"+cursoactual1+"'!AT$"+contador);
    hojactiva.getRange("G13").setFormula("='"+cursoactual1+"'!AU$"+contador);
    hojactiva.getRange("I13").setFormula("='"+cursoactual1+"'!AV$"+contador);
    hojactiva.getRange("K13").setFormula("='"+cursoactual1+"'!AW$"+contador);
    hojactiva.getRange("M13").setFormula("='"+cursoactual1+"'!AX$"+contador);
    hojactiva.getRange("O13").setFormula("='"+cursoactual1+"'!AY$"+contador);
     //FISICA
    hojactiva.getRange("C14").setFormula("='"+cursoactual1+"'!AZ$"+contador);
    hojactiva.getRange("E14").setFormula("='"+cursoactual1+"'!BA$"+contador);
    hojactiva.getRange("G14").setFormula("='"+cursoactual1+"'!BB$"+contador);
    hojactiva.getRange("I14").setFormula("='"+cursoactual1+"'!BC$"+contador);
    hojactiva.getRange("K14").setFormula("='"+cursoactual1+"'!BD$"+contador);
    hojactiva.getRange("M14").setFormula("='"+cursoactual1+"'!BE$"+contador);
    hojactiva.getRange("O14").setFormula("='"+cursoactual1+"'!BF$"+contador);

     //BIOLOGIA
    hojactiva.getRange("C15").setFormula("='"+cursoactual1+"'!BG$"+contador);
    hojactiva.getRange("E15").setFormula("='"+cursoactual1+"'!BH$"+contador);
    hojactiva.getRange("G15").setFormula("='"+cursoactual1+"'!BI$"+contador);
    hojactiva.getRange("I15").setFormula("='"+cursoactual1+"'!BJ$"+contador);
    hojactiva.getRange("K15").setFormula("='"+cursoactual1+"'!BK$"+contador);
    hojactiva.getRange("M15").setFormula("='"+cursoactual1+"'!BL$"+contador);
    hojactiva.getRange("O15").setFormula("='"+cursoactual1+"'!BM$"+contador);
    
    //ADMINISTRACION
    hojactiva.getRange("C16").setFormula("='"+cursoactual1+"'!BN$"+contador);
    hojactiva.getRange("E16").setFormula("='"+cursoactual1+"'!BO$"+contador);
    hojactiva.getRange("G16").setFormula("='"+cursoactual1+"'!BP$"+contador);
    hojactiva.getRange("I16").setFormula("='"+cursoactual1+"'!BQ$"+contador);
    hojactiva.getRange("K16").setFormula("='"+cursoactual1+"'!BR$"+contador);
    hojactiva.getRange("M16").setFormula("='"+cursoactual1+"'!BS$"+contador);
    hojactiva.getRange("O16").setFormula("='"+cursoactual1+"'!BT$"+contador);
     //SIC
    hojactiva.getRange("C17").setFormula("='"+cursoactual1+"'!BU$"+contador);
    hojactiva.getRange("E17").setFormula("='"+cursoactual1+"'!BV$"+contador);
    hojactiva.getRange("G17").setFormula("='"+cursoactual1+"'!BW$"+contador);
    hojactiva.getRange("I17").setFormula("='"+cursoactual1+"'!BX$"+contador);
    hojactiva.getRange("K17").setFormula("='"+cursoactual1+"'!BY$"+contador);
    hojactiva.getRange("M17").setFormula("='"+cursoactual1+"'!BZ$"+contador);
    hojactiva.getRange("O17").setFormula("='"+cursoactual1+"'!CA$"+contador);
     //INFORMATICA APLICADA
    hojactiva.getRange("C18").setFormula("='"+cursoactual1+"'!CB$"+contador);
    hojactiva.getRange("E18").setFormula("='"+cursoactual1+"'!CC$"+contador);
    hojactiva.getRange("G18").setFormula("='"+cursoactual1+"'!CE$"+contador);
    hojactiva.getRange("I18").setFormula("='"+cursoactual1+"'!CD$"+contador);
    hojactiva.getRange("K18").setFormula("='"+cursoactual1+"'!CF$"+contador);
    hojactiva.getRange("M18").setFormula("='"+cursoactual1+"'!CG$"+contador);
    hojactiva.getRange("O18").setFormula("='"+cursoactual1+"'!CH$"+contador);
    
    hojactiva.getRange("C19").setFormula("='"+cursoactual1+"'!CI$"+contador);
    hojactiva.getRange("E19").setFormula("='"+cursoactual1+"'!CJ$"+contador);
    hojactiva.getRange("G19").setFormula("='"+cursoactual1+"'!CK$"+contador);
    hojactiva.getRange("I19").setFormula("='"+cursoactual1+"'!CL$"+contador);
    hojactiva.getRange("K19").setFormula("='"+cursoactual1+"'!CM$"+contador);
    hojactiva.getRange("M19").setFormula("='"+cursoactual1+"'!CN$"+contador);
    hojactiva.getRange("O19").setFormula("='"+cursoactual1+"'!CO$"+contador);
    
    hojactiva.getRange("C20").setFormula("='"+cursoactual1+"'!CP$"+contador);
    hojactiva.getRange("E20").setFormula("='"+cursoactual1+"'!CQ$"+contador);
    hojactiva.getRange("G20").setFormula("='"+cursoactual1+"'!CR$"+contador);
    hojactiva.getRange("I20").setFormula("='"+cursoactual1+"'!CS$"+contador);
    hojactiva.getRange("K20").setFormula("='"+cursoactual1+"'!CT$"+contador);
    hojactiva.getRange("M20").setFormula("='"+cursoactual1+"'!CU$"+contador);
    hojactiva.getRange("O20").setFormula("='"+cursoactual1+"'!CV$"+contador);
  }
  else
  {
    if(curso == cursoactual2)
    {
    //MATEMATICA
    hojactiva.getRange("C7").setFormula("='"+cursoactual2+"'!C$"+contador);
    hojactiva.getRange("E7").setFormula("='"+cursoactual2+"'!D$"+contador);
    hojactiva.getRange("G7").setFormula("='"+cursoactual2+"'!E$"+contador);
    hojactiva.getRange("I7").setFormula("='"+cursoactual2+"'!F$"+contador);
    hojactiva.getRange("K7").setFormula("='"+cursoactual2+"'!G$"+contador);
    hojactiva.getRange("M7").setFormula("='"+cursoactual2+"'!H$"+contador);
    hojactiva.getRange("O7").setFormula("='"+cursoactual2+"'!I$"+contador);
    //lENGUA
    hojactiva.getRange("C8").setFormula("='"+cursoactual2+"'!j$"+contador);
    hojactiva.getRange("E8").setFormula("='"+cursoactual2+"'!K$"+contador);
    hojactiva.getRange("G8").setFormula("='"+cursoactual2+"'!L$"+contador);
    hojactiva.getRange("I8").setFormula("='"+cursoactual2+"'!M$"+contador);
    hojactiva.getRange("K8").setFormula("='"+cursoactual2+"'!N$"+contador);
    hojactiva.getRange("M8").setFormula("='"+cursoactual2+"'!O$"+contador);
    hojactiva.getRange("O8").setFormula("='"+cursoactual2+"'!P$"+contador);
     //INGLES
    hojactiva.getRange("C9").setFormula("='"+cursoactual2+"'!Q$"+contador);
    hojactiva.getRange("E9").setFormula("='"+cursoactual2+"'!R$"+contador);
    hojactiva.getRange("G9").setFormula("='"+cursoactual2+"'!S$"+contador);
    hojactiva.getRange("I9").setFormula("='"+cursoactual2+"'!T$"+contador);
    hojactiva.getRange("K9").setFormula("='"+cursoactual2+"'!U$"+contador);
    hojactiva.getRange("M9").setFormula("='"+cursoactual2+"'!V$"+contador);
    hojactiva.getRange("O9").setFormula("='"+cursoactual2+"'!W$"+contador);
     //EDUCACION FISICA
    hojactiva.getRange("C10").setFormula("='"+cursoactual2+"'!X$"+contador);
    hojactiva.getRange("E10").setFormula("='"+cursoactual2+"'!Y$"+contador);
    hojactiva.getRange("G10").setFormula("='"+cursoactual2+"'!Z$"+contador);
    hojactiva.getRange("I10").setFormula("='"+cursoactual2+"'!AA$"+contador);
    hojactiva.getRange("K10").setFormula("='"+cursoactual2+"'!AB$"+contador);
    hojactiva.getRange("M10").setFormula("='"+cursoactual2+"'!AC$"+contador);
    hojactiva.getRange("O10").setFormula("='"+cursoactual2+"'!AD$"+contador);
     //FEC
    hojactiva.getRange("C11").setFormula("='"+cursoactual2+"'!AE"+contador);
    hojactiva.getRange("E11").setFormula("='"+cursoactual2+"'!AF$"+contador);
    hojactiva.getRange("G11").setFormula("='"+cursoactual2+"'!AG"+contador);
    hojactiva.getRange("I11").setFormula("='"+cursoactual2+"'!AH"+contador);
    hojactiva.getRange("K11").setFormula("='"+cursoactual2+"'!AI"+contador);
    hojactiva.getRange("M11").setFormula("='"+cursoactual2+"'!AJ"+contador);
    hojactiva.getRange("O11").setFormula("='"+cursoactual2+"'!AK$"+contador);
     //HISTORIA
    hojactiva.getRange("C12").setFormula("='"+cursoactual2+"'!AL$"+contador);
    hojactiva.getRange("E12").setFormula("='"+cursoactual2+"'!AM$"+contador);
    hojactiva.getRange("G12").setFormula("='"+cursoactual2+"'!AN$"+contador);
    hojactiva.getRange("I12").setFormula("='"+cursoactual2+"'!AO$"+contador);
    hojactiva.getRange("K12").setFormula("='"+cursoactual2+"'!AP$"+contador);
    hojactiva.getRange("M12").setFormula("='"+cursoactual2+"'!AQ$"+contador);
    hojactiva.getRange("O12").setFormula("='"+cursoactual2+"'!AR$"+contador);
     //GEOGRAFIA
    hojactiva.getRange("C13").setFormula("='"+cursoactual2+"'!AS$"+contador);
    hojactiva.getRange("E13").setFormula("='"+cursoactual2+"'!AT$"+contador);
    hojactiva.getRange("G13").setFormula("='"+cursoactual2+"'!AU$"+contador);
    hojactiva.getRange("I13").setFormula("='"+cursoactual2+"'!AV$"+contador);
    hojactiva.getRange("K13").setFormula("='"+cursoactual2+"'!AW$"+contador);
    hojactiva.getRange("M13").setFormula("='"+cursoactual2+"'!AX$"+contador);
    hojactiva.getRange("O13").setFormula("='"+cursoactual2+"'!AY$"+contador);
     //QUIMICA
    hojactiva.getRange("C14").setFormula("='"+cursoactual2+"'!AZ$"+contador);
    hojactiva.getRange("E14").setFormula("='"+cursoactual2+"'!BA$"+contador);
    hojactiva.getRange("G14").setFormula("='"+cursoactual2+"'!BB$"+contador);
    hojactiva.getRange("I14").setFormula("='"+cursoactual2+"'!BC$"+contador);
    hojactiva.getRange("K14").setFormula("='"+cursoactual2+"'!BD$"+contador);
    hojactiva.getRange("M14").setFormula("='"+cursoactual2+"'!BE$"+contador);
    hojactiva.getRange("O14").setFormula("='"+cursoactual2+"'!BF$"+contador);

     //BIOLOGIA
    hojactiva.getRange("C15").setFormula("='"+cursoactual2+"'!BG$"+contador);
    hojactiva.getRange("E15").setFormula("='"+cursoactual2+"'!BH$"+contador);
    hojactiva.getRange("G15").setFormula("='"+cursoactual2+"'!BI$"+contador);
    hojactiva.getRange("I15").setFormula("='"+cursoactual2+"'!BJ$"+contador);
    hojactiva.getRange("K15").setFormula("='"+cursoactual2+"'!BK$"+contador);
    hojactiva.getRange("M15").setFormula("='"+cursoactual2+"'!BL$"+contador);
    hojactiva.getRange("O15").setFormula("='"+cursoactual2+"'!BM$"+contador);

   //ADMINISTRACION
    hojactiva.getRange("C16").setFormula("='"+cursoactual2+"'!BN$"+contador);
    hojactiva.getRange("E16").setFormula("='"+cursoactual2+"'!BO$"+contador);
    hojactiva.getRange("G16").setFormula("='"+cursoactual2+"'!BP$"+contador);
    hojactiva.getRange("I16").setFormula("='"+cursoactual2+"'!BQ$"+contador);
    hojactiva.getRange("K16").setFormula("='"+cursoactual2+"'!BR$"+contador);
    hojactiva.getRange("M16").setFormula("='"+cursoactual2+"'!BS$"+contador);
    hojactiva.getRange("O16").setFormula("='"+cursoactual2+"'!BT$"+contador);
     //SIC
    hojactiva.getRange("C17").setFormula("='"+cursoactual2+"'!BU$"+contador);
    hojactiva.getRange("E17").setFormula("='"+cursoactual2+"'!BV$"+contador);
    hojactiva.getRange("G17").setFormula("='"+cursoactual2+"'!BW$"+contador);
    hojactiva.getRange("I17").setFormula("='"+cursoactual2+"'!BX$"+contador);
    hojactiva.getRange("K17").setFormula("='"+cursoactual2+"'!BY$"+contador);
    hojactiva.getRange("M17").setFormula("='"+cursoactual2+"'!BZ$"+contador);
    hojactiva.getRange("O17").setFormula("='"+cursoactual2+"'!CA$"+contador);
     //INFORMATICA APLICADA
    hojactiva.getRange("C18").setFormula("='"+cursoactual2+"'!CB$"+contador);
    hojactiva.getRange("E18").setFormula("='"+cursoactual2+"'!CC$"+contador);
    hojactiva.getRange("G18").setFormula("='"+cursoactual2+"'!CD$"+contador);
    hojactiva.getRange("I18").setFormula("='"+cursoactual2+"'!CE$"+contador);
    hojactiva.getRange("K18").setFormula("='"+cursoactual2+"'!CF$"+contador);
    hojactiva.getRange("M18").setFormula("='"+cursoactual2+"'!CG$"+contador);
    hojactiva.getRange("O18").setFormula("='"+cursoactual2+"'!CH$"+contador);
    
    
    hojactiva.getRange("C19").setFormula("='"+cursoactual2+"'!CI$"+contador);
    hojactiva.getRange("E19").setFormula("='"+cursoactual2+"'!CJ$"+contador);
    hojactiva.getRange("G19").setFormula("='"+cursoactual2+"'!CK$"+contador);
    hojactiva.getRange("I19").setFormula("='"+cursoactual2+"'!CL$"+contador);
    hojactiva.getRange("K19").setFormula("='"+cursoactual2+"'!CM$"+contador);
    hojactiva.getRange("M19").setFormula("='"+cursoactual2+"'!CN$"+contador);
    hojactiva.getRange("O19").setFormula("='"+cursoactual2+"'!CO$"+contador);
    
    
    hojactiva.getRange("C20").setFormula("='"+cursoactual2+"'!CP$"+contador);
    hojactiva.getRange("E20").setFormula("='"+cursoactual2+"'!CQ$"+contador);
    hojactiva.getRange("G20").setFormula("='"+cursoactual2+"'!CR$"+contador);
    hojactiva.getRange("I20").setFormula("='"+cursoactual2+"'!CS$"+contador);
    hojactiva.getRange("K20").setFormula("='"+cursoactual2+"'!CT$"+contador);
    hojactiva.getRange("M20").setFormula("='"+cursoactual2+"'!CU$"+contador);
    hojactiva.getRange("O20").setFormula("='"+cursoactual2+"'!CV$"+contador);
    }
  }
  
}


function crearAlumno(){
  var fullactiu = boletin.getActiveSheet();
  var nombrefulls= boletin.getNumSheets();
  boletin.setActiveSheet(boletin.getSheets()[nombrefulls-1]);//Activa la última hoja para insertar el nuevo al final
  var alumnonuevo= boletin.insertSheet(); 
  alumnonuevo.insertColumnAfter(1);
  alumnonuevo.insertColumnAfter(1);
  alumnonuevo.insertColumnAfter(1);
  var notaalumno = boletin.getSheets()[boletin.getNumSheets()-1];
  notaalumno.getRange("A1").setValue("CICLO LECTIVO");
  notaalumno.getRange("A1").setFontWeight("bold");
  notaalumno.getRange("A1").setFontFamily("Roboto");
  notaalumno.getRange("A1").setHorizontalAlignment("center");
  notaalumno.setColumnWidth(1,100);
  

}
function test(nombre,seccion,contador){
   var totalSheet = boletin.getNumSheets();
   var templateSheet = boletin.getSheetByName("Template")
   boletin.insertSheet(nombre,totalSheet, {template:templateSheet})
   var alumno= boletin.getActiveSheet()
   alumno.getRange('B2').setValue(nombre);
   alumno.getRange('B3').setValue(seccion)
   contadorAlumnos(contador)
   setMaterias()
   setNotas()
   protegerHojas();
}
function buscaPromedios(){
  var prom =boletin.setActiveSheet(boletin.getSheetByName("PROMEDIO"));
  var hojas =boletin.getSheets();
  var nota;
  //var prom = boletin.getSheets()[boletin.getSheetByName("PROMEDIO")];
  var nombre;
  var cont=3
  for( var i=2;i<boletin.getNumSheets()-1;i++){
    nota = hojas[i].getRange("G3").getValue();
    nombre = hojas[i].getRange("B2").getValue();
    if(nota>6){
       prom.getRange("A"+cont).setValue(nombre);
       prom.getRange("B"+cont).setValue(nota);  
       cont++;
    }
  }

}
function protegerHojas(){
  var sheet= boletin.getActiveSheet()
  var proteccion =sheet.protect().setDescription("Bloqueo alumnos")
  var me = Session.getActiveUser()
  proteccion.removeEditors(proteccion.getEditors());
        if (proteccion.canDomainEdit()) {
             proteccion.setDomainEdit(false);
           }
              proteccion.addEditor(me)
              proteccion.addEditor('secundario@institutoaustro.edu.ar')
  }
function generarBoletines(){
  var cursoSelect = boletin.getActiveSheet();
  var curso = cursoSelect.getName();
  var alumnos = cursoSelect.getRange("B7:B30").getValues();
  var contador =cursoSelect.getRange("A50").getValues();
  for (var i = 0 ;i<alumnos.length;i++)
  {
    if(alumnos[i]!=""){
      test(alumnos[i],curso,i);
    }
  }
 

}


function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('page')
      .setWidth(860)
      .setHeight(500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Proteger Materias');
}

function getLotsOfThings(){
  var list = boletin.getActiveSheet().getRange("C4:CP4").getValues();//Obtiene las materias de la hoja activa
  var materias=[]
  var subArray=list[0]
  for(var i=0 ; i<subArray.length;i++){
    if(subArray[i]!=""){
        materias.push(subArray[i])
      }
  }      
  return materias
}
function obtenerProfes(){
  //var lista = boletin.getSheets()[3].getRange("C5:BQ5").getValues();
  var lista = boletin.getActiveSheet().getRange("C5:CP5").getValues();
  var listaProfes=[]
  var subArray=lista[0]
  for(var i =0;i<subArray.length-1;i=i+7){
    if(subArray[i]!=""){
      listaProfes.push(subArray[i])
    }
    else{
      listaProfes.push("S/profesor")
    }
  }
  return listaProfes
}
function getCorreos(elemento){
  var profes= obtenerProfes()
  var correoData = boletin.getSheets()[0].getRange("A3:B36").getValues();
  var listaCorreos=[]
  for(var i=0;i<correoData.length;i++){
    var usuario = correoData[i]
      if(usuario[0]===elemento)
      {
        var correo = usuario[1]
      }
      
  
  }
 
  return correo
}
function crearObjeto(){
  var materias = getLotsOfThings()
  var profes = obtenerProfes()
  var rangos = getRangos()
  //var correo =obtenerCorreo()
  var data = new Object()
  var arrayData = []
  for (var i=0;i<materias.length;i++){
    var data = new Object()
    var rango = rangos[i]
    //data.materia = materias[i]
    data.materias = materias[i]
    data.profes=profes[i]
    data.correo = getCorreos(profes[i]);
    if (rango[0]===materias[i]){
      data.rango = rango[1]
    }
    else{
      data.rango ="ne"
    }
    arrayData.push(data)
  }
  return arrayData
}
function getRangos(){
  var ss =boletin.getSheets()[0].getRange("F3:G30").getValues();
  return ss
}
function proteger(){
  var ss = SpreadsheetApp.getActive()
  var rangosMaterias = getRangos()
  var data = crearObjeto();
  var flag = false;
  var cont=3
  for (var  i = 0 ; i<rangosMaterias.length;i++){
    var rangos=rangosMaterias[i]
    for(var j = 0;j<data.length;j++){
      if(data[j].materias === rangos[0]){
        var rango = data[i].rango
        var description = data[i].materias
        if(data[i].correo !=""){
          var correo = data[i].correo
        } 
       // boletin.getSheets()[1].getRange("H"+cont).setValue("active")
        ss.addEditor(correo)
        var proteccion = ss.getRange(rango).protect().setDescription(description)
        
        proteccion.removeEditors(proteccion.getEditors());
        if (proteccion.canDomainEdit()) {
           proteccion.setDomainEdit(false);
         }
         proteccion.addEditor(correo)
         proteccion.addEditor("secundario@institutoaustro.edu.ar")
         flag = true
    }  
      
    }
    cont++
  }
  

  return flag
}
function eliminar(){
 var ss = SpreadsheetApp.getActive();
 var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
 for (var i = 0; i < protections.length; i++) {
   var protection = protections[i];
   if (protection.canEdit()) {
     protection.remove();
   }
 }

}
function hojas(){
  var nhojas= boletin.getSheets()
  var total =boletin.getNumSheets()
  
  var cursoactual1=nhojas[0].getName();//obtengo el nombre de la hoja del primer curso
  var cursoactual2=nhojas[2].getName();//obtengo el nombre de la hoja de segundo curso
  cursoactual2
}