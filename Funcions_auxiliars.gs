function esborradB (){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Formulari',""); 
};

function sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}