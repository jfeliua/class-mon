function onInstall(e) {
  onOpen(e)
};

/**
 * Afegeix el menú al full
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  if (e && e.authMode === ScriptApp.AuthMode.NONE){
    switch(Session.getActiveUserLocale()){
      case "ca":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activar el complement CLASS-MON','activaCLASSMON')
        .addToUi()
        break;      
      case "es":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activar el complemento CLASS-MON','activaCLASSMON')
        .addToUi()
         break;
      case "fr":
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Activez CLASS-MON','activaCLASSMON')
        .addToUi()
        break;
      default:
        SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Enable CLASS-MON','activaCLASSMON')
        .addToUi()
    };
  }else{
    var properties = PropertiesService.getDocumentProperties();
    var importacio = properties.getProperty('Importacio');
    if (importacio===null){
      importacio= "0";
      properties.setProperty('Importacio', "0");
    };
    var idioma = properties.getProperty('Idioma');
    if (idioma===null){
      idioma= Session.getActiveUserLocale();
      properties.setProperty('Idioma', idioma);
    }
    var formulari = properties.getProperty('Formulari');
    if (formulari===null){
      formulari= "0";
      properties.setProperty('Formulari', "0");
    };
    switch(idioma){
      case "ca":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Crear la plantilla CLASS-MON')
                      .addItem('Activitats','creaCLASSMON')
                      .addItem('Actituds','creaCLASSMON_ACT'))
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Canviar idioma')
                      .addItem('Español', 'espanol')
                      .addItem('English', 'english')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          if (formulari != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Crear el formulari','creaformulari_plantilla')
            .addSeparator()
            .addItem('Importar alumnes de Google Classroom','impalClasroom')
            .addToUi()
          }else{   
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Formulari')
                        .addItem('Obtenir l\'enllaç del formulari','enllaFormulari')
                        .addItem('Enviar el formulari als alumnes','enviaFormulari')
                        .addItem('Publicar l\'enllaç del formulari a Classroom com un anunci','classFormulari')
                        .addSeparator()
                        .addItem('Actualitzar el formulari amb noves activitats/sessions','actualitza_form')
                        .addSeparator()
                        .addItem('Torna a crear el formulari','creaformulari_plantilla'))
            .addSeparator()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Enllaç pels alumnes')
                        .addItem('Crear un enllaç web per a cada alumne per veure les respostes','fulls_alumnes')
                        .addItem('Enviar l\'enllaç als alumnes','enviaEnlla'))
             .addSubMenu(SpreadsheetApp.getUi().createMenu('Respostes')
                        .addItem('Recuperar respostes del formulari','proRespostes'))
            .addSeparator()
            .addItem('Importar alumnes de Google Classroom','impalClasroom')
            .addSeparator()
            .addToUi()
          }
        };
        break;
      case "es":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Crear la plantilla CLASS-MON')
                      .addItem('Actividades','creaCLASSMON')
                      .addItem('Actitudes','creaCLASSMON_ACT'))
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Cambiar idioma')
                      .addItem('Català', 'catala')
                      .addItem('English', 'english')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          if (formulari != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Crear el formulario','creaformulari_plantilla')
            .addSeparator()
            .addItem('Importar alumnos de Google Classroom','impalClasroom') 
            .addToUi()
          }else{   
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Formulario')
                        .addItem('Obtener el enlace del formulario','enllaFormulari')
                        .addItem('Enviar el formulario a los alumnos','enviaFormulari')
                        .addItem('Publicar el enlace del formulario en Classroom como un anuncio','classFormulari')
                        .addSeparator()
                        .addItem('Actualizar el formulario con nuevas actividades/sesiones','actualitza_form')
                        .addSeparator()
                        .addItem('Volver a crear el formulario','creaformulari_plantilla'))                        
            .addSeparator()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Enlace de los alumnos')
                        .addItem('Crear un enlace web para cada alumno para visualizar las respuestas','fulls_alumnes')
                        .addItem('Enviar el enlace a los alumnos','enviaEnlla'))
             .addSubMenu(SpreadsheetApp.getUi().createMenu('Respuestas')
                        .addItem('Recuperar las respuestas del formulario','proRespostes'))
            .addSeparator()
            .addItem('Importar alumnos de Google Classroom','impalClasroom') 
            .addSeparator()
            .addToUi()
          }
        };
        break;
      case "fr":
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Créez un gabarit de CLASS-MON')
                      .addItem('Activités','creaCLASSMON')
                      .addItem('Attitudes','creaCLASSMON_ACT'))
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Changez de langue')
                      .addItem('Català', 'catala')
                      .addItem('Español', 'espanol')
                      .addItem('English', 'english'))
          .addToUi()
        }else{ 
          if (formulari != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Créez le formulaire','creaformulari_plantilla')
            .addSeparator()
            .addItem('Importez élèves et enseignants de Google Classroom','impalClasroom')
            .addToUi()
          }else{   
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Formulaire')
                        .addItem('Obtenez le lien au formulairee','enllaFormulari')
                        .addItem('Envoyez le formulaire aux élèves','enviaFormulari')
                        .addItem('Publiez le lien au formulaire dans Classroom comme annonce','classFormulari')
                        .addSeparator()
                        .addItem('Mettre à jour le formulaire avec de nouvelles activités/sessions','actualitza_form')
                        .addSeparator()
                        .addItem('Recréez le formulaire','creaformulari_plantilla'))
            .addSeparator()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Lien étudiant')
                        .addItem('Créez un lien Web pour chaque élève afin qu\'il puisse consulter les réponses.','fulls_alumnes')
                        .addItem('Envoyer le lien aux étudiants','enviaEnlla'))
             .addSubMenu(SpreadsheetApp.getUi().createMenu('Responses')
                        .addItem('Récupérer les réponses au formulaire','proRespostes'))
            .addSeparator()
            .addItem('Importez élèves et enseignants de Google Classroom','impalClasroom')
            .addSeparator()
            .addToUi()
          }
        };
        break;
      default:
        if (importacio != "1") {
          SpreadsheetApp.getUi()
          .createAddonMenu()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Create template of CLASS-MON')
                      .addItem('Activities','creaCLASSMON')
                      .addItem('Attitudes','creaCLASSMON_ACT'))
          .addSeparator()
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Change language')
                      .addItem('Català', 'catala')
                      .addItem('Español', 'espanol')
                      .addItem ('Français', 'français'))
          .addToUi()
        }else{ 
          if (formulari != "1") {
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addItem('Create the form','creaformulari_plantilla')
            .addSeparator()
            .addItem('Import students from Google Classroom','impalClasroom')
            .addToUi()
          }else{   
            SpreadsheetApp.getUi()
            .createAddonMenu()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Form')
                        .addItem('Get the form link','enllaFormulari')
                        .addItem('Send form to students','enviaFormulari')
                        .addItem('Publish the form link in Classroom like an annoucement','classFormulari')
                        .addSeparator()
                        .addItem('Update the form with new activities/sessions','actualitza_form')
                        .addSeparator()
                        .addItem('Recreate the form','creaformulari_plantilla'))
            .addSeparator()
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Students links')
                        .addItem('Create a web link for each student to view the answers','fulls_alumnes')
                        .addItem('Send link to students','enviaEnlla'))
             .addSubMenu(SpreadsheetApp.getUi().createMenu('Responses')
                        .addItem('Retrieve form responses','proRespostes'))
            .addSeparator()
            .addItem('Import students from Google Classroom','impalClasroom')
            .addSeparator()
            .addToUi()
          }
        };
    };
  };
};

function activaCLASSMON(){
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nouformulari = Browser.msgBox('CLASS-MON','CLASS-MON s\'ha activat correctament', Browser.Buttons.OK);
      break;
    case "es":
      var nouformulari = Browser.msgBox('CLASS-MON','CLASS-MON se ha activado correctamente', Browser.Buttons.OK);
      break;
    case "fr":
      var nouformulari = Browser.msgBox('CLASS-MON','CLASS-MON est activé', Browser.Buttons.OK);
      break;
    default: 
      var nouformulari = Browser.msgBox('CLASS-MON','CLASS-MON is enabled', Browser.Buttons.OK);
  };
  onOpen();
};

function creaCLASSMON_ACT(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Tipus', "act"); 
  creaCLASSMON();
};

function creaCLASSMON(){
  var properties = PropertiesService.getDocumentProperties(); 
  var idioma = properties.getProperty('Idioma');   
  var classmon_tipus = properties.getProperty('Tipus');
  if (classmon_tipus!='act'){
    classmon_tipus='tasq';
  };
  switch(idioma){
    case "ca":
      var avis = Browser.msgBox('Crear CLASS-MON','Aquest procés eliminarà tots els fulls de càlcul existents i crearà la plantilla CLASS-MON. Voleu continuar?', Browser.Buttons.YES_NO);
      break;
    case "es":
      var avis = Browser.msgBox('Crear CLASS-MON','Este proceso eliminará todas las hojas de cálculo existentes y creará la plantilla CLASS-MON. ¿Continuar?', Browser.Buttons.YES_NO);
      break;
    case "fr":
      var avis = Browser.msgBox('Créez CLASS-MON','Ce processus supprimera toutes les feuilles de la feuille de calcul et créera un gabarit CLASS-MON.  Vous voulez continuer?', Browser.Buttons.YES_NO);
      break;
    default:
      var avis = Browser.msgBox('Create CLASS-MON','This process deletes all sheets in this spreadsheet and will create a CLASS-MON template. Do you wish to continue?', Browser.Buttons.YES_NO);
  }  
  if (avis==='yes'){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var nom_full_alumnes= "Alumnes";
        var nom_full_activitats_2 = "Activitats";
        if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
        }else{
          var nom_full_activitats = "Activitats";
        };
        var nom_sessio = "Sessió";
        var ajuda_avaluacio = "Puntueu l'actitud de cada sessió entre 0 i 3";
        var nom_full_avaluacio= "Avaluació";
        var nom_full_respostes= "Respostes_num";
        var nom_full_resp_img= "Respostes";
        var nom_full_imatges= "Imatges";
        var nom_full_monitor="Monitor-alumnes";
        var nom_actitud= "Actitud";
        var nom_full_monitor_act_2 = "Monitor-activitats";
        if (classmon_tipus==='act'){
          var nom_full_monitor_act = "Monitor-actituds";
        }else{
          var nom_full_monitor_act = "Monitor-activitats";
        };
        var nom_full_dif="Dif";
        var nom_full_comentaris="Comentaris";
        var nom_full_fitxa="Fitxa alumne";
        var source = SpreadsheetApp.openById("1DxcOgRtkBWVfWzgWgnKIkPgHuOCrBhZASrLhj3FabtA");
        if (classmon_tipus==='act'){
          var desc1 ="He estat distret pràcticament tota l'hora, sense fer les tasques ni escoltar les explicacions.";
          var desc2 ="En molts moment m'he distret amb els companys i no he realitzat totes les tasques.";
          var desc3="He estat atent i treballant força estona, tot i que en algun moment m'he distret.";
          var desc4="He aprofitat molt bé tota la sessió, escoltant i treballant quan calia.";
        }         
        break;
      case "es":
        var nom_full_alumnes= "Alumnos";
        var nom_full_activitats_2 = "Actividades";
        if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
        }else{
          var nom_full_activitats = "Actividades";
        };
        var nom_sessio = "Sesión";
        var ajuda_avaluacio = "Puntuar la actitud en cada sesión entre 0 y 3";
        var nom_full_avaluacio= "Evaluación";
        var nom_full_respostes= "Respuestas_num";
        var nom_full_resp_img= "Respuestas";
        var nom_full_imatges= "Imágenes";
        var nom_full_monitor="Monitor-alumnos";
        var nom_actitud= "Actitud";
        var nom_full_monitor_act_2 = "Monitor-actividades";
        if (classmon_tipus==='act'){
          var nom_full_monitor_act = "Monitor-actitudes";
        }else{
          var nom_full_monitor_act = "Monitor-actividades";
        };
        var nom_full_dif="Dif";
        var nom_full_comentaris="Comentarios";
        var nom_full_fitxa="Ficha alumno";
        var source = SpreadsheetApp.openById("1x_o-lzdMgo0i4SIGWaZ2aKWvRCkmdjlOV4SOxaNR9cY");
        if (classmon_tipus==='act'){
          var desc1 ="He estado distraído prácticamente toda la hora, sin hacer las tareas ni escuchar las explicaciones.";
          var desc2 ="En muchos momento me he distraído con los compañeros y no he realizado todas las tareas.";
          var desc3="He estado atento y trabajando bastante rato, aunque en algún momento me he distraído.";
          var desc4="He aprovechado muy bien toda la sesión, escuchando y trabajando cuando era necesario.";
        } 
        break;
      case "fr":
        var nom_full_alumnes= "Élèves";
        var nom_full_activitats_2 = "Activitéss";
        if (classmon_tipus==='act'){
          var nom_full_activitats = "Séances";
        }else{
          var nom_full_activitats = "Activitéss";
        };
        var nom_sessio = "Séance";
        var ajuda_avaluacio = "Notez l'attitude dans chaque session entre 0 et 3";
        var nom_full_avaluacio= "Évaluation";
        var nom_full_respostes= "Résponses_num";
        var nom_full_resp_img= "Résponses";
        var nom_full_imatges= "Images";
        var nom_full_monitor="Monitor-élèves";
        var nom_actitud= "Attitude";
        var nom_full_monitor_act_2 = "Monitor-activités";
        if (classmon_tipus==='act'){
          var nom_full_monitor_act = "Monitor-attitudes";
        }else{
          var nom_full_monitor_act = "Monitor-activités";
        };
        var nom_full_dif="Dif";
        var nom_full_comentaris="Observations";
        var nom_full_fitxa="Dossier de l'élève";
        var source = SpreadsheetApp.openById("1gPB_nHKVVZi6pKoYER6AbGtBA7880pG6uS9WOwgPkeI");
        if (classmon_tipus==='act'){
          var desc1 ="J'ai été pratiquement distraite tout le temps, ne faisant pas mes devoirs ou n'écoutant pas les explications.";
          var desc2 ="À plusieurs reprises, j'ai été distrait par mes collègues et je n'ai pas accompli toutes les tâches.";
          var desc3="J'ai été attentif et j'ai travaillé pas mal de temps, bien qu'à un moment donné j'ai été distrait.";
          var desc4="J'ai fait bon usage de toute la séance, en écoutant et en travaillant quand c'était nécessaire.";
        } 
        break;
      default:
        var nom_full_alumnes= "Students";
        var nom_full_activitats_2 = "Activities";
        if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
        }else{
          var nom_full_activitats = "Activities";
        };
        var nom_sessio = "Session";
        var ajuda_avaluacio = "Rate the attitude in each session between 0 and 3";
        var nom_full_avaluacio= "Evaluation";
        var nom_full_respostes= "Responses_num";
        var nom_full_resp_img= "Responses";
        var nom_full_imatges= "Pictures";
        var nom_full_monitor="Monitor-Students";
        var nom_actitud= "Attitude";
        var nom_full_monitor_act_2 = "Monitor-Activities";
        if (classmon_tipus==='act'){
          var nom_full_monitor_act = "Monitor-attitudes";
        }else{
          var nom_full_monitor_act = "Monitor-Activities";
        };
        var nom_full_dif="Dif";
        var nom_full_comentaris="Comments";
        var nom_full_fitxa="Student's file";
        var source = SpreadsheetApp.openById("1P0DqnwZ7zJj8OW5vz1vPSfHrchT4wxr9ZWPmNW9eZYI");
        if (classmon_tipus==='act'){
          var desc1 ="I've been practically distracted all the time, not doing work or listening to explanations.";
          var desc2 ="At many times I was distracted by my colleagues and did not complete all the tasks.";
          var desc3="I've been attentive and working quite a while, although at some point I've been distracted.";
          var desc4="I made good use of the whole session, listening and working when necessary.";
        } 
    }  
    var sheet = source.getSheetByName(nom_full_activitats_2);
    var sheet1 = source.getSheetByName(nom_full_alumnes);
    var sheet2 = source.getSheetByName(nom_full_avaluacio);
    var sheet5 = source.getSheetByName(nom_full_resp_img);
    var sheet3 = source.getSheetByName(nom_full_respostes);
    var sheet4 = source.getSheetByName(nom_full_imatges);
    var sheet6 = source.getSheetByName(nom_full_monitor);
    var sheet7 = source.getSheetByName(nom_full_dif);
    var sheet8 = source.getSheetByName(nom_full_monitor_act_2);
    var sheet9 = source.getSheetByName(nom_full_comentaris);
    var sheet10 = source.getSheetByName(nom_full_fitxa);
    var destination = SpreadsheetApp.getActiveSpreadsheet();
    for (var i=0; i<destination.getNumSheets();i++){
      var full1 = destination.getSheets()[i].setName("Full" + i);
    };
    var full_act=sheet.copyTo(destination);
    full_act.setName(nom_full_activitats);
    if (classmon_tipus==='act'){
      full_act.getRange(1,2,1,1).setValue(nom_sessio);
    };
    var full_al = sheet1.copyTo(destination); 
    full_al.setName(nom_full_alumnes);
    var full_resp = sheet3.copyTo(destination); 
    full_resp.setName(nom_full_respostes);
    if (classmon_tipus==='act'){
      full_resp.getRange(1,3,1,1).setFormula("transpose('"+ nom_full_activitats +"'!A2:A81");
    };
    var full_img = sheet4.copyTo(destination); 
    full_img.setName(nom_full_imatges);
    if (classmon_tipus==='act'){
      full_img.getRange(2,1,1,1).setValue(desc1);
      full_img.getRange(3,1,1,1).setValue(desc2);
      full_img.getRange(4,1,1,1).setValue(desc3);
      full_img.getRange(5,1,1,1).setValue(desc4);
    };    
    var full_resp_img = sheet5.copyTo(destination); 
    full_resp_img.setName(nom_full_resp_img);
    var full_aval = sheet2.copyTo(destination);
    full_aval.setName(nom_full_avaluacio);
    if (classmon_tipus==='act'){
      full_aval.getRange(1,3,1,1).setValue(ajuda_avaluacio);
      full_aval.getRange(2,3,1,1).setFormula("transpose('"+ nom_full_activitats +"'!A2:A81)");
    };
    var full_dif = sheet7.copyTo(destination); 
    full_dif.setName(nom_full_dif);
    if (classmon_tipus==='act'){
      full_dif.getRange(2,3,1,1).setFormula("transpose('"+ nom_full_activitats +"'!A2:A81)");
    };
    var full_mon = sheet6.copyTo(destination); 
    full_mon.setName(nom_full_monitor);
    if (classmon_tipus==='act'){
      full_mon.getRange(1,3,1,1).setValue(nom_actitud);
      full_mon.getRange(2,3,1,1).setFormula("transpose('"+ nom_full_activitats +"'!A2:A81)");
    };
    var full_fitxa = sheet10.copyTo(destination); 
    full_fitxa.setName(nom_full_fitxa);
    if (classmon_tipus==='act'){
      full_fitxa.getRange(2,1,1,1).setFormula("={"+ nom_full_activitats +"!B:C}");
    };
    var full_mon_act = sheet8.copyTo(destination);
    full_mon_act.setName(nom_full_monitor_act);
    if (classmon_tipus==='act'){
      full_mon_act.getRange(2,1,1,1).setValue(nom_actitud);
      full_mon_act.getRange(4,1,1,1).setFormula("={"+ nom_full_activitats +"!A2:C81}");
    };
    var full_comentaris = sheet9.copyTo(destination); 
    full_comentaris.setName(nom_full_comentaris);
    if (classmon_tipus==='act'){
      full_comentaris.getRange(1,3,1,1).setFormula("transpose('"+ nom_full_activitats +"'!A2:A81)");
    };
    sleep (2000);
    var destination = SpreadsheetApp.getActiveSpreadsheet();
    var fulls= destination.getNumSheets();
    for (var i=0; i<fulls-11;i++){
      var full1 = destination.getSheetByName("Full" + i)
      destination.deleteSheet(full1);
    };
    full_resp.hideSheet();
    full_img.hideSheet();
    full_dif.hideSheet();
    full_resp_img.hideSheet();
    full_comentaris.hideSheet();
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('Importacio', "1");
    documentProperties.setProperty('Formulari', "0");
    
    //Canviar el menú, treient Crear CLASS-MON i posant el que correspongui
    esborradB();
    onOpen();
  };
};

function catala(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "ca"); 
  onOpen();
};

function espanol(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "es"); 
  onOpen();
};

function english(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "en"); 
  onOpen();
};

function français(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Idioma', "fr"); 
  onOpen();
};

function impalClasroom(){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  switch(idioma){
    case "ca":
      var nom_html='importal_ca';
      break;
    case "es":
      var nom_html='importal_es';
      break;
    case "fr":
      var nom_html='importal_fr';
      break;
    default:
      var nom_html='importal_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON');
};

function importacio_al(formObject){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma'); 
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON');
  var cursid = formObject.combo_curs;
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      var nom_html='Cal triar un curs de Classroom';
      var curs_m='Curs';
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      var nom_html='Es necesario elegir un curso de Classroom';
      var curs_m='Curso';
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      var nom_html='Il est nécessaire de choisir un cours de Classroom';
      var curs_m='Cours';
      break;
    default:
      var nom_full_alumnes= "Students";
      var nom_html='It is necessary to choose a Classroom course';
      var curs_m='Course';
  }; 
  if (cursid == 0){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    impalClasroom();
  }else{
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('cursid', cursid);
    //importem alumnes
    var pagina=null;
    var ki=0;
    var estudiants=0;
    var alumnes = [];
    do {
      alumnes[ki]=Classroom.Courses.Students.list(cursid,{pageToken:pagina});  //Classroom treu els alumnes de 30 en 30. Cal llegir 30 i després canviar el token per llegir-ne 30 més
      estudiants=estudiants + alumnes[ki].students.length;
      var pagina=alumnes[ki].nextPageToken;
      ki++;
    }while (pagina);
    var matriu=new Array(estudiants);
    var comptador=0;
    for (var f=0;f<alumnes.length;f++){
      for (var i=0;i<alumnes[f].students.length;i++){
        var cognom_al=alumnes[f].students[i].profile.name.familyName;
        var nom_al=alumnes[f].students[i].profile.name.givenName;
        var mail_al=alumnes[f].students[i].profile.emailAddress;
        matriu[comptador]=new Array(2);
        matriu[comptador][0]=cognom_al+", "+nom_al;
        matriu[comptador][1]=mail_al;
        comptador++;
      };
    };
    
    matriu.sort();
    var rang_full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_full_alumnes).getRange(2,1,estudiants,2);
    rang_full.setValues(matriu);
    
    switch(idioma){
      case "ca":
        var frase = properties.setProperty('frase', 'Els alumnes s\'han importat correctament');
        var boto = properties.setProperty('boto', 'Tancar finestra');
        break;
      case "es":
        var frase = properties.setProperty('frase', 'Los alumnos se han importado correctamente');
        var boto = properties.setProperty('boto', 'Cerrar ventana');
        break;    
      case "fr":
        var frase = properties.setProperty('frase', 'Élèves ont été importés correctement');
        var boto = properties.setProperty('boto', 'Fermez');
        break;
      default:
        var frase = properties.setProperty('frase', 'Students have been properly imported');
        var boto = properties.setProperty('boto', 'Close');
    };  
    var nom_html='confirma';
    var html = HtmlService
    .createTemplateFromFile(nom_html)
    .evaluate();
    
    SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON'); 
  };   
};

function creaformulari_plantilla(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  var classmon_tipus = properties.getProperty('Tipus');
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
      }else{
          var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1cFDvfwPVkk8KLZ60XpjZtGPkSLBXI2megory2oZEWhU/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/16pwc_LVH8_x6x1f1SXONph8ij4up3kYC_IVQ-laDyz0/edit";
      };
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
      }else{
          var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1JaP4q2hjFqflt0B4h2Qr452aankN5Q_m-gtI0K7_umc/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/1n1up7rOAcQjEHR9lhsq98_KHRYEL_NJ5Rdv035c2vCI/edit";
      };
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1XadNO2-o86Ox4wZ2rkuJFn3DKDCgr9dGHTIFe-HT2lk/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/15T1horzWaOBAe43whc8xX2BB5yB-gjfDPm20yJLuTfk/edit";
      };
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1ma1vJCDiqiiovKKNonY3GJib2ZSluVt2WIBWAL-wl6I/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/1kP6ebr7IQAmsNaM1agCjGGlZmKPWoxi-_t7u3NmNVHw/edit";
      };
  }  
  
  //Comprovem que al full hi hagi Acivitats introduïdes
  var Activitats = llibreActual.getSheetByName(nom_full_activitats);
  var rangactivitats = Activitats.getDataRange();
  var rad = rangactivitats.getValues()
  var n_act=0;
  //Mirem quantes activitats tenen títol
  for (var i=1;i<rangactivitats.getNumRows();i++){
    if (rad[i][1]!=""){
      ++n_act;
    };
  };  
  var Alumness = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = Alumness.getDataRange();
  var acts = [];
  var desc = []
  var nombreacts=0;
  if (n_act===0){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        if (classmon_tipus==='act'){
          Browser.msgBox('Sessions','No has indicat cap sessió!', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activitats','No has indicat cap activitat!', Browser.Buttons.OK);
        };
        break;
      case "es":
        if (classmon_tipus==='act'){
          Browser.msgBox('Sesiones','¡No has indicado ninguna sesión!', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Actividades','¡No has indicado ninguna actividad!', Browser.Buttons.OK);
        };        
        break;
      case "fr":
        if (classmon_tipus==='act'){
          Browser.msgBox('Séances','La liste des séances est vide', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activités','La liste d\'activités est vide', Browser.Buttons.OK);
        };
        break;
      default:
        if (classmon_tipus==='act'){
          Browser.msgBox('Sessions','The list of sessions is empty.', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activities','The list of activities is empty.', Browser.Buttons.OK);
        };
    } 
    return;
  }
  
  //Eliminem el format del full d'activitats  
  var rangesborrar=Activitats.getRange(2,2,rangactivitats.getNumRows()-1,rangactivitats.getNumColumns()-1);
  rangesborrar.clearFormat();
  rangesborrar.setVerticalAlignment('middle');
  rangesborrar.setHorizontalAlignment('left')
  rangesborrar.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangesborrar.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  //Preguntem pel nom del formulari
  switch(idioma){
    case "ca":
      var nomform = Browser.inputBox('Nom del formulari','Quin nom vols que tingui el formulari?', Browser.Buttons.OK_CANCEL);
      if (classmon_tipus==='act'){
        var descform = "Formulari per al seguiment del treball a classe i de l'actitud. Quan acabi la sessió, indica com has estat a l'aula."
        var t_activitat="Sessió?";
      }else{
        var descform = "Formulari per al seguiment de les tasques. Quan acabis una tasca, indica com l'has realitzat."
        var t_activitat="Activitat?";
      };
      break;
    case "es":
      var nomform = Browser.inputBox('Nombre del formulario','¿Qué nombre quieres que tenga el formulario?', Browser.Buttons.OK_CANCEL);
      if (classmon_tipus==='act'){
        var descform = "Formulario para el seguimiento del trabajo en clase y de la actitud. Cuando termine la sesión, indica como ha sido tu actitud en clase."
        var t_activitat="¿Sesión?";
      }else{
        var descform = "Formulario para el seguimiento de las tareas. Cuando termines cada tarea, indica como la has realizado."
        var t_activitat="¿Actividad?";
      };
      break;
    case "fr":
      var nomform = Browser.inputBox('Nom du formulaire','Quel est le nom du formulaire?', Browser.Buttons.OK_CANCEL);
      if (classmon_tipus==='act'){
        var descform = "Formulaire de suivi du travail en classe et de l'attitude. A la fin de la session, indiquez votre attitude en classe.";
        var t_activitat="Séance?";
      }else{
        var descform = "Formulaire de suivi des tâches. Lorsque vous avez terminé chaque tâche, indiquez comment vous l'avez exécutée."
        var t_activitat="Activité?";
      };
      break;
    default:
      var nomform = Browser.inputBox('Form name','What is the name of the form?', Browser.Buttons.OK_CANCEL);
      if (classmon_tipus==='act'){
      var descform = "Form for monitoring work in class and attitude. At the end of the session, indicate your attitude in class."
      var t_activitat="Session?";
      }else{
        var descform = "Form to follow up the tasks. When you finish each task, indicate how you have performed it."
        var t_activitat="Activity?";
      };
  } 
  //Fem copia del formulari plantilla
  var form_pl = FormApp.openByUrl(form_plantilla);
  var form_pl_id= form_pl.getId()
  var form_dr = DriveApp.getFileById(form_pl_id).makeCopy(nomform);
  var formid = form_dr.getId();
  var form = FormApp.openById(formid);
  form.setCollectEmail(true); 
  form.setLimitOneResponsePerUser(false);
  form.setAllowResponseEdits(false);
  form.setTitle(nomform);
  form.setDescription(descform);  
  try{ form.setRequireLogin(true);} catch (error) {};
  var formurl = form.getPublishedUrl();
  
  //Movem el formulari creat a la carpeta del full de CLASS-MON
  var folders = DriveApp.getFileById(llibreActual.getId()).getParents(); //Agafem les carpetes on està el full de CoRubrics
  var folder = folders.next();  //Agafem la primer carpeta.  
  if (folder.getName()!=DriveApp.getRootFolder().getName()){  //Si es troba a La Meva Unitat, no movem el formulari
    var fitxer = DriveApp.getFileById(formid);
    folder.addFile(fitxer);
    DriveApp.removeFile(fitxer);    
  };
  
  //Deso l'ID en un ScriptdB   
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('Formid',formid);
  documentProperties.setProperty('Formurl',formurl);
  documentProperties.setProperty('Formnom',nomform);
  documentProperties.setProperty('Formulari',"1");
  documentProperties.setProperty('Activitats',n_act);
  
  //Agafem el nom de les activitats i la seva descripció (corregint possibles espais al final)
  for (var i=0; i<rangactivitats.getNumRows()-1;i++){
    var comp = rangactivitats.getCell(i+2,2).getValue();
    if (comp===""){
      break;
    };
    acts[i] = comp;
    desc[i] = rangactivitats.getCell(i+2,3).getValue();
    acts[i] = acts[i].toString();
    desc[i] = desc[i].toString();
    var fora_espais_finals = acts[i].trim();
    if (acts[i] != fora_espais_finals){ //eliminem possibles espais al final del nom de l'activitat
      rangactivitats.getCell(i+2,2).setValue(fora_espais_finals);
    };
    acts[i]=fora_espais_finals;
    var fora_espais_finals = desc[i].trim();
    if (desc[i] != fora_espais_finals){ //eliminem possibles espais al final de la descripció de l'activitat
      rangactivitats.getCell(i+2,3).setValue(fora_espais_finals);
    };
    desc[i]=fora_espais_finals; 
    
  };    

  //Creem la primera pregunta (més endavant omplirem les opcions)
  var pregunta1List =form.getItems(FormApp.ItemType.LIST)[0];
  pregunta1List.setTitle(t_activitat);
  var preguntaList=pregunta1List.asListItem();
  var pregid = preguntaList.getId();
  var urlpersonalitzat = 'ARRAYFORMULA(IF(B2:B="";"";HYPERLINK("' + formurl + '?usp=pp_url&entry.1395216507="&B2:B)))';
  
  var preg= form.getItems(FormApp.ItemType.MULTIPLE_CHOICE); //Agafem la primera pregunta de la plantilla (amb imatges)
  preg[0].setTitle(acts[0])
  var obs=form.getItems(FormApp.ItemType.PARAGRAPH_TEXT);
  
  var sec= form.getItems(FormApp.ItemType.PAGE_BREAK); //Agafem la primera secció i posem el títol i la descripció correctes
  sec[0].setTitle(acts[0]);
  sec[0].setHelpText(desc[0]);

  //Creem les seccions (duplicant la primera) i actualitzo la pregunta inicial de l'activitat a valorar
  var seccio = [];
  seccio[0]=sec[0];                    
  for (i=1; i<acts.length;i++){ //per cada activitat duplico la secció
    seccio[i] = sec[0].duplicate();
    seccio[i].setTitle(acts[i]);
    seccio[i].setHelpText(desc[i]);
    seccio[i].setGoToPage(FormApp.PageNavigationType.SUBMIT);
    var pregunta = preg[0].duplicate();
    pregunta.setTitle(acts[i])
    obs[0].duplicate();
  }  
  
  //Omplim la primera pregunta (Activitat?) i fem que salti segons la resposta
  var llista_opcions = []; 
  var seccio= form.getItems(FormApp.ItemType.PAGE_BREAK);
  for (i=0; i<acts.length;i++){ //afegeixo cada activitat a la pregunta inicial
    llista_opcions[i] = preguntaList.createChoice(acts[i],seccio[i].asPageBreakItem());
  };
  preguntaList.setChoices(llista_opcions);
  
  //Indiquem l'enllaç per omplert en el full Activitats
  llibreActual.getSheetByName(nom_full_activitats).getRange("D2").setFormula(urlpersonalitzat);  
  
  //Fem que el matexi full sigui el full destí de respostes
  form.setDestination(FormApp.DestinationType.SPREADSHEET, llibreActual.getId());
  sleep (2000);
  //var full_form=SpreadsheetApp.openById(llibreActual.getId()).getSheets()[0];
  //full_form.hideSheet();

  try{
    //Elimino qualsevol trigger anterior
    var allTriggers = ScriptApp.getProjectTriggers(); //Si ja haviem creat un script, l'eliminem
    for (var i = 0; i < allTriggers.length; i++) {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
    //Deso al full d'estadístiques que s'ha creat el formualari
    var cv="https://docs.google.com/spreadsheets/d/1fScPpy4FqbxyTDguH-ORH_Hw_DgpTxOSrJRIQb72i5s/";
    var fullOrigen = SpreadsheetApp.openByUrl(cv).getSheetByName("Analytics");
    var filesple = fullOrigen.getDataRange().getNumRows()+1;
    var range = fullOrigen.getRange("A" + filesple + ":B" + filesple);
    var avui = new Date();
    var data_actual11 = avui.getDate(); //Trobo dia d'avui
    var data_actual = new Date();
    data_actual.setDate(data_actual11);
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        range.setValues([["CLASS-MON ca",data_actual]]);
        break;
      case "es":
        range.setValues([["CLASS-MON es",data_actual]]);
        break;
      case "fr":
        range.setValues([["CLASS-MON fr",data_actual]]);
        break;
      default:
        range.setValues([["CLASS-MON en",data_actual]]);
    }
    //Creo el nou disparador
    var sheet = SpreadsheetApp.getActive();
    ScriptApp.newTrigger("proRespostes")
    .forSpreadsheet(sheet)
    .onFormSubmit()
    .create();        
  }  
  catch(err){
    Logger.log(err);
  };
    
  onOpen();
}

/**
 * Mostra l'enllaç del formulari
 * per pantalla
 */
function enllaFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();  
  var idioma = properties.getProperty('Idioma');  
  var classmon_tipus = properties.getProperty('Tipus');
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
      }else{
          var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
  }   
  //Recupero el ID del fomrulari, del ScriptDB
  var documentProperties = PropertiesService.getDocumentProperties();
  var formurl= documentProperties.getProperty('Formurl');
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      var enl = '<p>L\'enllaç del formulari és: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    case "es":
      var enl = '<p>El enlace del formulario es: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    case "fr":
      var enl = '<p>Le lien au formulaire est:: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
      break;
    default:
      var enl = '<p>The link to the form is: </p><p><a target="_blank" href="'+formurl+'">'+formurl+'</a></p>';
  }  
  
  var htmlApp = HtmlService.createHtmlOutput();
  htmlApp.setContent(enl);
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
      htmlApp.setTitle('Enllaç del formulari');
      break;
    case "es":
      htmlApp.setTitle('Enlace del formulario');
      break;
    case "fr":
        htmlApp.setTitle('Lien au formulaire');
      break;
    default:
      htmlApp.setTitle('Form link');
  } 
  htmlApp.setWidth(400);
  htmlApp.setHeight(150);

  SpreadsheetApp.getActive().show(htmlApp);
  
};

/*
 * A partir de les respostes dels alumnes,
 * omple el full de càlcul Avaluació.
 * Si un alumne marca dos cops la mateixa pregunta, només s'agafa la última
 */
function proRespostes() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();  
  var idioma = properties.getProperty('Idioma');   
  var classmon_tipus = properties.getProperty('Tipus'); 
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
      }else{
          var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      var nom_full_monitor="Monitor-alumnes";
      var nom_full_comentaris="Comentaris";
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
      }else{
          var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      var nom_full_monitor="Monitor-alumnos";
      var nom_full_comentaris="Comentarios";
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      var nom_full_monitor="Monitor-élèves";
      var nom_full_comentaris="Observations";
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
      var nom_full_monitor="Monitor-Students";
      var nom_full_comentaris="Comments";
  }  
  var respostesActual = llibreActual.getSheetByName(nom_full_respostes);
  var monitorActual = llibreActual.getSheetByName(nom_full_comentaris);
  var rangrespostes = respostesActual.getDataRange();
  var alumnesActual = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = alumnesActual.getDataRange();
  if (rangalumnes.getNumRows()>1){ //si no ha entrat alumnes, no fem res i mostrem error.
    //Recupero el ID del formulari i el nombre d'activitats
    var formid = properties.getProperty('Formid');
    var form = FormApp.openById(formid);
    var num_act = parseInt(properties.getProperty('Activitats'));
    //Recupero les respostes del formulari
    var re = form.getResponses();
    var num_resp = re.length;
    var resposta_al= []; //Defineixo una matriu bidimensional per acumular les respotes d'un mateix alumne. En cada fila, una resposta del formulari
    var comentaris_al=[];//Matriu pels comentaris del alumnes
    var matriu_final=[] //Matriu amb nombre enlloc de respostes i sense mails
    var nombre_alumnes=rangalumnes.getNumRows()-1;  //Busquem quants alumnes hi ha 
    for (var i=0;i<nombre_alumnes;i++) {  //Defineixo la matriu com a bidimensional i l'omplo amb els alumnes.
      resposta_al[i]=[];
      comentaris_al[i]=[];
      matriu_final[i]=[];
      for (var j=0; j<num_act+1;j++){
        if (j===0){
          resposta_al[i][j]=rangalumnes.getValues()[i+1][1];
        }else{
          resposta_al[i][j]="";
          comentaris_al[i][j-1]="";
          matriu_final[i][j-1]="";
        };
      };
    };
    //Mirem resposta per resposta del formulari
    for (var i=0; i < num_resp; i++) {
      var resposta_f = re[i];
      var mail_alum=resposta_f.getRespondentEmail(); //Recuperem qui ha respost
      var alumne=-1; //variable per trobar la fila de la matriu on està l'alumne
      //Cerco l'alumne a la matriu i li afegim la resposta (si un alumne contesta dues vegades la mateix resposta, només comptem la segona)
      for (var k=0; k<nombre_alumnes; k++){
        if (resposta_al[k][0]===mail_alum){
          alumne=k;
        };
      };
      if (alumne!=-1){ //si contesta un adreça que no pertany a cap alumne, ho ignorem.
        var itemResponses = resposta_f.getItemResponses(); //D'aquesta resposta, recuperem les respostes a cada pregunta 
        for (var j=1;j<itemResponses.length;j++) { //Per cada pregunta,recuperem la resposta i la desem a la matriu per acumular respostes
          var itemResponse = itemResponses[j]; 
          var preg_respo=itemResponse.getItem().getIndex(); //Com que a cada resposta només es respon a una pregunta, recuperem quina pregunta ha respost
          resposta_al[alumne][(preg_respo+1)/3]=itemResponse.getResponse(); 
          j++;
          if(j<itemResponses.length){
            itemResponse = itemResponses[j]; 
            comentaris_al[alumne][((preg_respo+1)/3)-1]=itemResponse.getResponse();
          }else{
            comentaris_al[alumne][((preg_respo+1)/3)-1]="";
          };
        };
      };
    };
    //Substituïm les respostes pels números que hi ha al full Imatges i eliminem el mail dels alumnes
    var full_imatges = llibreActual.getSheetByName(nom_full_imatges);
    var rangimatges = full_imatges.getDataRange();
    var dades_imatges = rangimatges.getValues();
    for (i=0; i<nombre_alumnes; i++){
      for (j=1; j<num_act+1;j++){
        for (k=1;k<5;k++){
          if (resposta_al[i][j]===dades_imatges[k][0]){
            matriu_final[i][j-1]=dades_imatges[k][1];
          };
        };
      };
    };
    //Creem una matriu sense el mail dels alumnes
    respostesActual.getRange(2,3,nombre_alumnes,num_act).setValues(matriu_final); //Copiem el resultat al full Respostes 
    monitorActual.getRange(2,3,nombre_alumnes,num_act).setValues(comentaris_al); //Copiem els comentaris com a notes
  }else{ //si no ha introduit alumnes mostrem error 
    switch(idioma){
      case "ca":
        var miss_al = '<p>Has d\'indicar com a mínim un alumne</p>';
        break;
      case "es":
        var miss_al = '<p>Debes indicar al menos un alumno</p>';
        break;
      case "fr":
        var miss_al = '<p>Vous devez indiquer au moins un étudiant</p>';
        break;
      default:
        var miss_al = '<pYou must indicate at least one student</p>';
    } 
    var htmlApp = HtmlService.createHtmlOutput();
    htmlApp.setContent(miss_al);
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        htmlApp.setTitle('Alumnes');
        break;
      case "es":
        htmlApp.setTitle('Alumnos');
        break;
      case "fr":
        htmlApp.setTitle('Étudiants');
        break;
      default:
        htmlApp.setTitle('Student');
    } 
    htmlApp.setWidth(400);
    htmlApp.setHeight(150);
    
    SpreadsheetApp.getActive().show(htmlApp);
    
  };
}


/**
 * Envia l'enllaç del formulari per mail
 * a tots les alumnes.
 */
function enviaFormulari() {
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
   var classmon_tipus = properties.getProperty('Tipus');
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
      }else{
          var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
      }else{
          var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
  }  
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = llistaalumnes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var dades_alumnes = rangalumnes.getValues();
  
  //Recupero el ID del formulari, del ScriptDB
  var documentProperties = PropertiesService.getDocumentProperties();
  var formurl = documentProperties.getProperty('Formurl');
  var formnom = documentProperties.getProperty('Formnom');
  
  //Defineixo el títol (nom del formulari) i el cos del missatge
  var titolform = formnom;
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  switch(idioma){
    case "ca":
        var titol = "Formulari CLASS-MON: " + titolform;
        var cosmissatge = "Aquí teniu l'enllaç per indicar l'estat de les tasques: " + formurl;
      break;
    case "es":
        var titol = "Formulario  CLASS-MON: " + titolform;
        var cosmissatge = "Aquí teneis el enlace para indicar el estado de las tareas: " + formurl;
      break;
    case "fr":
        var titol = "Formulaire CLASS-MON: " + titolform;
        var cosmissatge = "Ici vous avez le lien pour indiquer l'état des tâches: " + formurl;
      break;
    default:
        var titol = "CLASS-MON Form: " + titolform;
        var cosmissatge = "Here you have the link to indicate the state of the tasks: " + formurl;
  } 
  
  //Envio el formulari a cada un dels alumnes
  var alumnes = "";
  for (var i=1; i<nombrealumnes+1;i++){
    alumnes = dades_alumnes[i][1];
    if (alumnes!=""){
      GmailApp.sendEmail(alumnes, titol, cosmissatge);
      alumnes="";
    };
  };  
};

function classFormulari(){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  switch(idioma){
    case "ca":
      var nom_html='clasF_ca';
      break;
    case "es":
      var nom_html='clasF_es';
      break;
    case "fr":
      var nom_html='clasF_fr';
      break;
    default:
      var nom_html='clasF_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();  
  SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON');  
};

function classform(formObject){
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma'); 
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON');
  var cursid = formObject.combo_curs;
  var titol = formObject.titol;
  var descripcio = formObject.descripcio;
  switch(idioma){
    case "ca":
      var nom_html='Cal triar un curs de Classroom';
      var curs_m='Curs';
      break;
    case "es":
      var nom_html='Es necesario elegir un curso de Classroom';
      var curs_m='Curso';
      break;
    case "fr":
      var nom_html='Classroomeko ikasgela hautatzea beharrezkoa da';
      var curs_m='Curso';
      break;
    default:
      var nom_html='It is necessary to choose a Classroom course';
      var curs_m='Course';
  }; 
  if (cursid == 0){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    classFormulari();
  }else{  
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('cursid', cursid);
    var fomrid = documentProperties.getProperty('Formid');
    var formurl = documentProperties.getProperty('Formurl');
    //CREAR L'ANUNCI AMB EL FORMULARI ADJUNT
    var creo_anunci = {
      "courseId": cursid,
      "text": titol,
      'materials': [  
        {'link': { 'url': formurl }} 
      ],  
      "state": "PUBLISHED"
    }
    var anunci_creat=Classroom.Courses.Announcements.create(creo_anunci, cursid)
    var anunci_id=anunci_creat.id;
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('anunci_id', anunci_id);  
    switch(idioma){
      case "ca":
        var frase = properties.setProperty('frase', 'El formulari s\'ha publicat correctament al Classroom');
        var boto = properties.setProperty('boto', 'Tancar finestra');
        break;
      case "es":
        var frase = properties.setProperty('frase', 'El formulario se ha publicado correctamente en Classroom');
        var boto = properties.setProperty('boto', 'Cerrar ventana');
        break;    
      case "fr":
        var frase = properties.setProperty('frase', 'Le formulaire a été publié dans Classroom avec succès');
        var boto = properties.setProperty('boto', 'Fermez');
        break;
      default:
        var frase = properties.setProperty('frase', 'The form has been successfully published in Classroom');
        var boto = properties.setProperty('boto', 'Close');
    };  
    var nom_html='confirma';
    var html = HtmlService
    .createTemplateFromFile(nom_html)
    .evaluate();
    SpreadsheetApp.getUi().showModelessDialog(html, 'CLASS-MON');  
  };
};

//Crea un codi per alumne i crea l'enllaç a l'app web.

function fulls_alumnes(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();
  var idioma = properties.getProperty('Idioma'); 
  var classmon_tipus = properties.getProperty('Tipus'); 
  if (classmon_tipus!='act'){
    classmon_tipus='tasq';
  };
  //Modifico la adreça que passo com a paràmetre per tal que cap alumne l'utilitzi per veure el full class-mon
  var llibre_id=llibreActual.getId();
  var chars='0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var atzar=chars[Math.floor(Math.random() * chars.length)];
  var llibre_id_nou = llibre_id.substr(0, 1)+atzar+llibre_id.substr(1, llibre_id.length);
  llibre_id=llibre_id_nou;
  
  var enlla_full= "https://script.google.com/macros/s/AKfycbxkZA-exC6uox59yR1gwHQBSJ9MykwKMGmBSk707z7dDEjoCQ/exec?full="+llibre_id+"&lang="+idioma+"&tipus="+classmon_tipus+"&id=";   
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      var nom_full_class_mon_num="CLASS-MON_num";
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sesiones";
      }else{
        var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      var nom_full_class_mon_num="CLASS-MON_num";
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      var nom_full_class_mon_num="CLASS-MON_num";
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
      var nom_full_class_mon_num="CLASS-MON_num";
  }  
  var rangalumnes = llibreActual.getSheetByName(nom_full_alumnes).getDataRange()
  var dadesalumnes = rangalumnes.getValues(); 
  var result = [];
  for (var i=1;i<rangalumnes.getNumRows();i++){
    result[i-1]=[];
    result[i-1][0]='';
    
    var coincidencia=false;
    while (coincidencia===false){
      for (var j = 32; j > 0; --j){
        chars='0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
        result[i-1][0] += chars[Math.floor(Math.random() * chars.length)];
      };
      result[i-1][1]=enlla_full + result[i-1][0];
      //Comprovem que no inventi dos nombres iguals per a dos alumnes
      coincidencia=true;
      for (var k=0; k<(i-1);k++){
        if (result[i-1][0]===result[k][0]){
          coincidencia=false;          
        };        
      };
    };
  };
  llibreActual.getSheetByName(nom_full_alumnes).getRange(2,3,rangalumnes.getNumRows()-1,2).setValues(result);    
  //Compartim el full amb tothom que tingui el link per tal que la web app hi tingui accés
  var files = DriveApp.getFileById(llibreActual.getId());
  files.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
};


/**
 * Envia l'enllaç per mail
 * a tots les alumnes.
 */
function enviaEnlla(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');   
  var classmon_tipus = properties.getProperty('Tipus'); 
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sessions";
      }else{
          var nom_full_activitats = "Activitats";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      var m_error="Primer has de crea l'enllaç per als alumnes";
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
          var nom_full_activitats = "Sesiones";
      }else{
          var nom_full_activitats = "Actividades";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      var m_error="Primero debes crear el enlace para los alumnos";
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
      }else{
        var nom_full_activitats = "Activitéss";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      var m_error="Vous devez d'abord créer le lien pour les élèves";
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
      }else{
        var nom_full_activitats = "Activities";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
      var m_error="First you have to create the link for the students";
  }
  var llistaalumnes = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = llistaalumnes.getDataRange();
  var nombrealumnes = rangalumnes.getNumRows()-1;
  var dades_alumnes = rangalumnes.getValues();
  
  if (dades_alumnes[1][2]!=""){
    
    //Defineixo el títol (nom del formulari) i el cos del missatge
    var titolform = llibreActual.getName();
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        var titol = "Enllaç CLASS-MON: "+titolform;
        var cosmissatge = "Aquí teniu l'enllaç per veure l'estat de les tasques: ";
        break;
      case "es":
        var titol = "Enlace CLASS-MON: "+titolform;
        var cosmissatge = "Aquí teneis el enlace para ver el estado de las tareas: ";
        break;
      case "fr":
        var titol = "CLASS-MON Liason: "+titolform;
        var cosmissatge = "Ici vous avez le lien pour voir le statut des tâches: ";
        break;
      default:
        var titol = "CLASS-MON Link: "+titolform;
        var cosmissatge = "Here you have the link to see the state of the tasks: ";
    } 
    
    //Envio el formulari a cada un dels alumnes
    var alumnes = "";
    for (var i=1; i<nombrealumnes+1;i++){
      var cccmis=cosmissatge;
      alumnes = dades_alumnes[i][1];
      var enll = dades_alumnes[i][3];
      cccmis+=enll;
      if (alumnes!=""){
        GmailApp.sendEmail(alumnes, titol, cccmis);
        alumnes="";
      };
    };  
  }else{
    var msg=Browser.msgBox("CLASS-MON",m_error, Browser.Buttons.OK);       
  };
};


function actualitza_form(){
  var llibreActual = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties();   
  var idioma = properties.getProperty('Idioma');  
  var classmon_tipus = properties.getProperty('Tipus');
  switch(idioma){
    case "ca":
      var nom_full_alumnes= "Alumnes";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
        var t_activitat="Sessió?";
      }else{
        var nom_full_activitats = "Activitats";
        var t_activitat="Activitat?";
      };
      var nom_full_avaluacio= "Avaluació";
      var nom_full_respostes= "Respostes_num";
      var nom_full_imatges= "Imatges";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1cFDvfwPVkk8KLZ60XpjZtGPkSLBXI2megory2oZEWhU/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/16pwc_LVH8_x6x1f1SXONph8ij4up3kYC_IVQ-laDyz0/edit";
      };
      break;
    case "es":
      var nom_full_alumnes= "Alumnos";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sesiones";
        var t_activitat="¿Sesión?";
      }else{
        var nom_full_activitats = "Actividades";
        var t_activitat="¿Actividad?";
      };
      var nom_full_avaluacio= "Evaluación";
      var nom_full_respostes= "Respuestas_num";
      var nom_full_imatges= "Imágenes";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1JaP4q2hjFqflt0B4h2Qr452aankN5Q_m-gtI0K7_umc/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/1n1up7rOAcQjEHR9lhsq98_KHRYEL_NJ5Rdv035c2vCI/edit";
      };
      break;
    case "fr":
      var nom_full_alumnes= "Élèves";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Séances";
        var t_activitat="Séance?";
      }else{
        var nom_full_activitats = "Activitéss";
        var t_activitat="Activité?";
      };
      var nom_full_avaluacio= "Évaluation";
      var nom_full_respostes= "Résponses_num";
      var nom_full_imatges= "Images";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1XadNO2-o86Ox4wZ2rkuJFn3DKDCgr9dGHTIFe-HT2lk/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/15T1horzWaOBAe43whc8xX2BB5yB-gjfDPm20yJLuTfk/edit";
      };
      break;
    default:
      var nom_full_alumnes= "Students";
      if (classmon_tipus==='act'){
        var nom_full_activitats = "Sessions";
        var t_activitat="Session?";
      }else{
        var nom_full_activitats = "Activities";
        var t_activitat="Activity?";
      };
      var nom_full_avaluacio= "Evaluation";
      var nom_full_respostes= "Responses_num";
      var nom_full_imatges= "Pictures";
      if (classmon_tipus==='act'){
        var form_plantilla= "https://docs.google.com/forms/d/1ma1vJCDiqiiovKKNonY3GJib2ZSluVt2WIBWAL-wl6I/edit";
      }else{
        var form_plantilla= "https://docs.google.com/forms/d/1kP6ebr7IQAmsNaM1agCjGGlZmKPWoxi-_t7u3NmNVHw/edit";
      };
  }  
  
  //Comprovem que al full hi hagi Acivitats introduïdes
  var Activitats = llibreActual.getSheetByName(nom_full_activitats);
  var rangactivitats = Activitats.getDataRange();
  var rad = rangactivitats.getValues()
  var n_act=0;
  //Mirem quantes activitats tenen títol
  for (var i=1;i<rangactivitats.getNumRows();i++){
    if (rad[i][1]!=""){
      ++n_act;
    };
  };  
  var Alumness = llibreActual.getSheetByName(nom_full_alumnes);
  var rangalumnes = Alumness.getDataRange();
  var acts = [];
  var desc = []
  var nombreacts=0;
  if (n_act===0){
    var properties = PropertiesService.getDocumentProperties();   
    var idioma = properties.getProperty('Idioma');   
    switch(idioma){
      case "ca":
        if (classmon_tipus==='act'){
          Browser.msgBox('Sessions','No has indicat cap sessió!', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activitats','No has indicat cap activitat!', Browser.Buttons.OK);
        };
        break;
      case "es":
        if (classmon_tipus==='act'){
          Browser.msgBox('Sesiones','¡No has indicado ninguna sesión!', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Actividades','¡No has indicado ninguna actividad!', Browser.Buttons.OK);
        };        
        break;
      case "fr":
        if (classmon_tipus==='act'){
          Browser.msgBox('Séances','La liste des séances est vide', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activités','La liste d\'activités est vide', Browser.Buttons.OK);
        };
        break;
      default:
        if (classmon_tipus==='act'){
          Browser.msgBox('Sessions','The list of sessions is empty.', Browser.Buttons.OK);
        }else{
          Browser.msgBox('Activities','The list of activities is empty.', Browser.Buttons.OK);
        };
    } 
    return;
  }
  
  //Eliminem el format del full d'activitats  
  var rangesborrar=Activitats.getRange(2,2,rangactivitats.getNumRows()-1,rangactivitats.getNumColumns()-1);
  rangesborrar.clearFormat();
  rangesborrar.setVerticalAlignment('middle');
  rangesborrar.setHorizontalAlignment('left')
  rangesborrar.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  rangesborrar.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
 
  //Recuperem el formulari 
  var documentProperties = PropertiesService.getDocumentProperties();
  var formid = properties.getProperty('Formid');
  var form = FormApp.openById(formid);
  var n_act_original = parseInt(properties.getProperty('Activitats'));
  
  //Agafem el nom de les activitats i la seva descripció (corregint possibles espais al final)
  for (var i=0; i<rangactivitats.getNumRows()-1;i++){
    var comp = rangactivitats.getCell(i+2,2).getValue();
    if (comp===""){
      break;
    };
    acts[i] = comp;
    desc[i] = rangactivitats.getCell(i+2,3).getValue();
    acts[i] = acts[i].toString();
    desc[i] = desc[i].toString();
    var fora_espais_finals = acts[i].trim();
    if (acts[i] != fora_espais_finals){ //eliminem possibles espais al final del nom de l'activitat
      rangactivitats.getCell(i+2,2).setValue(fora_espais_finals);
    };
    acts[i]=fora_espais_finals;
    var fora_espais_finals = desc[i].trim();
    if (desc[i] != fora_espais_finals){ //eliminem possibles espais al final de la descripció de l'activitat
      rangactivitats.getCell(i+2,3).setValue(fora_espais_finals);
    };
    desc[i]=fora_espais_finals; 
    
  };    

  var preg= form.getItems(FormApp.ItemType.MULTIPLE_CHOICE); //Agafem la primera pregunta del formulari
  preg[0].setTitle(acts[0])
  var obs=form.getItems(FormApp.ItemType.PARAGRAPH_TEXT);
  
  var sec= form.getItems(FormApp.ItemType.PAGE_BREAK); //Agafem la primera secció i posem el títol i la descripció correctes
  sec[0].setTitle(acts[0]);
  sec[0].setHelpText(desc[0]);
  
  //Creem les seccions que faltin (duplicant la primera)
  var seccio = [];
  seccio[0]=sec[0];  
  if (n_act_original<n_act){ //Creem preguntes noves si n'ha afegit
    for (i=n_act_original; i<n_act;i++){ //per cada activitat duplico la secció
      seccio[i] = sec[0].duplicate();
      seccio[i].setTitle(acts[i]);
      seccio[i].setHelpText(desc[i]);
      seccio[i].setGoToPage(FormApp.PageNavigationType.SUBMIT);
      var pregunta = preg[0].duplicate();
      pregunta.setTitle(acts[i])
      obs[0].duplicate();
    }  
  }
  sec= form.getItems(FormApp.ItemType.PAGE_BREAK)
  if (n_act_original>n_act){ //Si no hi ha preguntes noves, eliminem les del final
    for (i=n_act; i<n_act_original;i++){ //per cada activitat duplico la secció
      Logger.log(n_act_original);
      form.deleteItem(sec[i]); //Eliminem la secció
      form.deleteItem(preg[i]); //Eliminem la pregunta de les caretes
      form.deleteItem(obs[i]); //Eliminem la pregunta de les observacions
    }
  };  
  sec= form.getItems(FormApp.ItemType.PAGE_BREAK); //Torno a agafar les seccions, que potser n'hem afegit
  preg= form.getItems(FormApp.ItemType.MULTIPLE_CHOICE); //Torno a agafar les preguntes de les caretes, que potser n'hem afegit
  for (i=1; i<n_act;i++){ //per cada activitat que ja hi havia, actualitzo el títol i la descripció
    seccio[i]=sec[i];
    seccio[i].setTitle(acts[i]);
    seccio[i].setHelpText(desc[i]);
    preg[i].setTitle(acts[i])
  }   
  properties.setProperty('Activitats',n_act); //Desem el nou nombre d'activitats  
  
 //Omplim la primera pregunta (Activitat?) i fem que salti segons la resposta
  var pregunta1List =form.getItems(FormApp.ItemType.LIST)[0];
  var preguntaList=pregunta1List.asListItem();
  var llista_opcions = []; 
  var seccio= form.getItems(FormApp.ItemType.PAGE_BREAK);
  for (i=0; i<acts.length;i++){ //afegeixo cada activitat a la pregunta inicial
    llista_opcions[i] = preguntaList.createChoice(acts[i],seccio[i].asPageBreakItem());
  };
  preguntaList.setChoices(llista_opcions);
    
  onOpen();

}
