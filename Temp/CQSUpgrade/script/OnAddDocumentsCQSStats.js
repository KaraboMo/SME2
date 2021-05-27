/* eslint-disable */
"use strict";
function OnAddDocuments(addType, sourcePath) {
  addTypeArray = ['Copy template', 'Copy/Paste or Drag/Drop from other WP file',
             'Copy/Paste or Drag/Drop from Explorer', 'Added via save as PDF',
             'Copy/Paste from same WP file', ' New document link'];
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("ADDDOCUMENT", addTypeArray[addType]);
    // app = null;
  }
  catch(e) {
  }
}
