function setProfesores() {
  var hojactiva=boletin.getActiveSheet();//Obtengo la hoja de calculo activa, el boletin
  var nhojas=boletin.getSheets();//obtengo la cantidad de hojas del boletin
  var cursoactual1=nhojas[0].getName();//obtengo el nombre de la hoja del primer curso
  var cursoactual2=nhojas[1].getName();//obtengo el nombre de la hoja de segundo curso
  var curso=hojactiva.getRange("B3").getValue();  //obtengo el valor del curso en la hoja del alumno
  if(curso ==cursoactual1)
  {
    hojactiva.getRange("A7").setFormula("='"+cursoactual1+"'!C$5");
    hojactiva.getRange("A8").setFormula("='"+cursoactual1+"'!I$5");
    hojactiva.getRange("A9").setFormula("='"+cursoactual1+"'!O$5");
    hojactiva.getRange("A11").setFormula("='"+cursoactual1+"'!AA$5");
    hojactiva.getRange("A12").setFormula("='"+cursoactual1+"'!AG$5");
    hojactiva.getRange("A13").setFormula("='"+cursoactual1+"'!AM$5");
    hojactiva.getRange("A14").setFormula("='"+cursoactual1+"'!AS$5");
    hojactiva.getRange("A15").setFormula("='"+cursoactual1+"'!AY$5");
    hojactiva.getRange("A16").setFormula("='"+cursoactual1+"'!BE$5");
    hojactiva.getRange("A17").setFormula("='"+cursoactual1+"'!BK$5");
    hojactiva.getRange("A18").setFormula("='"+cursoactual1+"'!BQ$5");
    /*hojactiva.getRange("A19").setFormula("='"+cursoactual1+"'!H$5);
    hojactiva.getRange("A19").setFormula("='"+cursoactual1+"'!H$5);*/
}
  else{
    if(curso == cursoactual2)
    {
      hojactiva.getRange("A7").setFormula("='"+cursoactual2+"'!C$5");
      hojactiva.getRange("A8").setFormula("='"+cursoactual2+"'!I$5");
      hojactiva.getRange("A9").setFormula("='"+cursoactual2+"'!O$5");
      hojactiva.getRange("A11").setFormula("='"+cursoactual2+"'!AA$5");
      hojactiva.getRange("A12").setFormula("='"+cursoactual2+"'!AG$5");
      hojactiva.getRange("A13").setFormula("='"+cursoactual2+"'!AM$5");
      hojactiva.getRange("A14").setFormula("='"+cursoactual2+"'!AS$5");
      hojactiva.getRange("A15").setFormula("='"+cursoactual2+"'!AY$5");
      hojactiva.getRange("A16").setFormula("='"+cursoactual2+"'!BE$5");
      hojactiva.getRange("A17").setFormula("='"+cursoactual2+"'!BK$5");
      hojactiva.getRange("A18").setFormula("='"+cursoactual2+"'!BQ$5");
     }
  
  }
}                                       
                                      
