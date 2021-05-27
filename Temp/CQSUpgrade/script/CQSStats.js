//CQS STATS JS
/* eslint-disable */
function CQSFileExists(sPathandFileName)
{
//This will return a boolean stating if the CW file is compressed or not
//The path and file name with the .ac should be passed in.
//if the ac file can be found false will return, otherwise true will return

 	//Check if logging or debugger has been turned on
	//Removing the call to avoid stack overflow
	try
    {
      var bResult = false;
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        if (oFSO.FileExists(sPathandFileName))
          bResult = true;
      }
                }
    catch(e)
    {
    }
  return bResult;
}

function CQSFolderExists(sPathName)
{
//This will return a boolean stating if the CW file is compressed or not
//The path and file name with the .ac should be passed in.
//if the ac file can be found false will return, otherwise true will return

	//Check if logging or debugger has been turned on
	//Removing the call to avoid stack overflow
	try
    {
      var bResult = false;
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        if (oFSO.FolderExists(sPathName))
          bResult = true;
      }
                }
    catch(e)
    {
    }
  return bResult;
}

function CQSGetFilePathLib(sPathandFileName)
{
//This will return the path of a specified file
//if c:\Program files\MyFile.txt was passed in
//it will return c:\Program files\
	//Check if logging or debugger has been turned on
	try
    {
      var sResult = "";
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        sResult = oFSO.GetParentFolderName(sPathandFileName)
      }
    }
    catch(e)
    {
    }
  return sResult;
}

var CQSStatsHTTP = (function() {
   var CQSWebService = "http://stats-wp.cqscloud.com/cwwp/";

   var XMLHttpRequest = function () {
     var oUtilities = new ActiveXObject("CaseWare.Utilities");
     return oUtilities.CreateXMLHttpRequest();
   };

   function getDomain() {
     var network = new ActiveXObject('WScript.Network');
     var userDomain = network.UserDomain;
     if(userDomain.toUpperCase() === "CQS" || userDomain.toUpperCase() === "ADAPTIT") {
       if(network.userName.toUpperCase() === "GEORGES"
	   || network.userName.toUpperCase() === "MADELEINEB") {
         CQSWebService = "http://localhost:3000/cwwp/"; //local server
        
       } else {
         CQSWebService = "http://cqs-stats-test.elasticbeanstalk.com/cwwp/"; //dev server
       }
     }
     network = null;
   }

   function send(statType, data) {
     getDomain();
     var r = new XMLHttpRequest();
     r.open("POST", CQSWebService + statType, false);
     r.setRequestHeader("Content-Type", "application/json");
     r.send(JSON.stringify(data));
   }

   function registeredTemplates(data) {
     send("templates", data);
   }

   function registerEvent(data) {
     send("files/EngFileAction", data);
   }

   function registerEngFileInfo(data) {
     send("files/EngFileInfo", data);
   }

   function RegisterUserEngFile(data) {
     send("files/RegisterUserEngFile", data);
   }


   return {
     registeredTemplates: registeredTemplates,
     registerEvent: registerEvent,
     registerEngFileInfo: registerEngFileInfo,
     RegisterUserEngFile: RegisterUserEngFile
   };
 });


var ComputerInformation = (function() {
  function getComputerInfoStructure() {
    return {
      computerName: "",
      userDomain: "",
      userName: ""
    };
  }

  function getCQSCloudInfo(compinfo) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var wsc = new ActiveXObject("WScript.Shell");
    var sFilePath = wsc.ExpandEnvironmentStrings("%ProgramFiles%") + "\\CQS\\nw\\app\\install\\user.json";
    compinfo.organisationCode = "";
    compinfo.emailAddress = "";
    var ForReading = 1;
    try {
      oFile = fso.OpenTextFile(sFilePath, ForReading);
      //get file contents
      sText = oFile.ReadAll();
      var UserObject = JSON.parse(sText);
      compinfo.organisationCode = UserObject.code;
      compinfo.emailAddress = UserObject.email;
    }
    catch(e) {}
    oFile = null;
    wsc = null;
    fso = null;
  }

  function getComputerInfo() {
    var network = new ActiveXObject('WScript.Network');
    var compinfo = getComputerInfoStructure();
    compinfo.computerName = network.computerName;
    compinfo.userDomain = network.UserDomain;
    compinfo.userName = network.Username;
    network = null;
    getCQSCloudInfo(compinfo);
    return compinfo;
  }

  return {
    getInfo: getComputerInfo
  };
});

var MetaDataInfo = (function () {

  function getFileMetaDataStructure() {
    return {
      Product: "",
      ProductLanguage: "",
      Country: "",
      Framework: "",
      Mapping: "",
      LibraryPath: "",
      AltProductLanguage: "",
      DateIssued: "",
      ExpM: "",
      ExpD: "",
      Alt_Version: "",
      Alt_Product: "",
      RelatedEntities: "",
      ScriptFunction: "",
      Alt_ScriptFunction: ""
    };
  }

  function getMFMAMetaDataStructure() {
    return {
      DocumentLibraryPath: "",
      DocumentLibraryFileName: "",
      GRAPDocumentLibraryPath: "",
      GRAPDocumentLibraryFileName: ""
    };
  }

  function getPROBEMetaDataStructure() {
    return {
      Probe_Language: "",
      Probe_Product: "",
      Probe_Country: "",
      Probe_Fversion: "",
      Probe_Cversion: "",
      Probe_Tversion: ""
    };
  }

  function getNTMetaDataStructure() {
    return {
      NT_Product: "",
      NT_Version: ""
    };
  }

  function getPFMANTMetaDataStructure() {
    return {
      PFMANT_Product: "",
      PFMANT_Version: ""
    };
  }

  function getTaxCompMetaDataStructure() {
    return {
      TC_Product: "",
      TC_Version: "",
      TC_ProductLanguage: "",
      TC_Country: "",
      TC_ProductWebLink: "",
      TC_EXPD: "",
      TC_ExpM: "",
      TC_DateIssued: "",
      TC_LibraryPath: "",
      TC_ScriptFunction: "",
      TC_DontAskForUpdate: ""
    };
  }

  function getMetaInfo(CaseWareFileName) {
      //ProductInformation
    var oCW = new ActiveXObject("CaseWare.Application");
    var l_Metadata = {};
    var l_TemplateMetaData = getFileMetaDataStructure();

    var metadataarray = oCW.Clients.GetMetaData(CaseWareFileName);
    var e = new Enumerator(metadataarray);
    prefix = "CWCustomProperty.",
      property = null,
      result = {};

    for (e.moveFirst(); !e.atEnd(); e.moveNext()) {
      property = e.item();
      //Only get custom properties.
      if (property.Name.indexOf(prefix) !== -1) {
        //Strip the CWCustomProperty prefix
        result[property.Name.substr(prefix.length)] = property.Value;
      }
    }
    //debugger;
    l_TemplateMetaData.Product = result["Product"];
    l_TemplateMetaData.ProductLanguage = result["ProductLanguage"];
    l_TemplateMetaData.Country = result["Country"];
    l_TemplateMetaData.Framework = result["Framework"]
    l_TemplateMetaData.Mapping = result["Mapping"];
    l_TemplateMetaData.LibraryPath = result["LibraryPath"];
    l_TemplateMetaData.AltProductLanguage = result["AltProductLanguage"];
    l_TemplateMetaData.DateIssued = result["DateIssued"];
    l_TemplateMetaData.ExpM = result["ExpM"];
    l_TemplateMetaData.ExpD = result["ExpD"];
    l_TemplateMetaData.Alt_Version = result["Alt_Version"];
    l_TemplateMetaData.Alt_Product = result["Alt_Product"];
    l_TemplateMetaData.RelatedEntities = result["RelatedEntities"];
    l_TemplateMetaData.ScriptFunction = result["ScriptFunction"];
    l_TemplateMetaData.Alt_ScriptFunction = result["Alt_ScriptFunction"];
    //MFMA Inforamtion
    l_MFMAMetaData = getMFMAMetaDataStructure();
    l_MFMAMetaData.DocumentLibraryPath = result["DocumentLibraryPath"];
    l_MFMAMetaData.DocumentLibraryFileName = result["DocumentLibraryFileName"];
    l_MFMAMetaData.GRAPDocumentLibraryPath = result["GRAPDocumentLibraryPath"];
    l_MFMAMetaData.GRAPDocumentLibraryFileName = result["GRAPDocumentLibraryFileName"];
    //getPROBEMetaDataStructure
    l_PROBEMetaData = getPROBEMetaDataStructure();
    l_PROBEMetaData.Probe_Language = result["Probe_Language"];
    l_PROBEMetaData.Probe_Product = result["Probe_Product"];
    l_PROBEMetaData.Probe_Country = result["Probe_Country"];
    l_PROBEMetaData.Probe_Fversion = result["Probe_Fversion"];
    l_PROBEMetaData.Probe_Cversion = result["Probe_Cversion"];
    l_PROBEMetaData.Probe_Tversion = result["Probe_Tversion"];
    //getNTMetaDataStructure
    l_NTMetaData = getNTMetaDataStructure();
    l_NTMetaData.NT_Product = result["NT_Product"];
    l_NTMetaData.NT_Version = result["NT_Version"];
    //getPFMANTMetaDataStructure
    l_PFMANTMetaData = getPFMANTMetaDataStructure();
    l_PFMANTMetaData.PFMANT_Product = result["PFMANT_Product"];
    l_PFMANTMetaData.PFMANT_Version = result["PFMANT_Version"];
    //getTaxCompMetaDataStructure
    l_TaxCompMetaData = getTaxCompMetaDataStructure();
    l_TaxCompMetaData.TC_Product = result["TC_Product"];
    l_TaxCompMetaData.TC_Version = result["TC_Version"];
    l_TaxCompMetaData.TC_ProductLanguage = result["TC_ProductLanguage"];
    l_TaxCompMetaData.TC_Country = result["TC_Country"];
    l_TaxCompMetaData.TC_ProductWebLink = result["TC_ProductWebLink"];
    l_TaxCompMetaData.TC_EXPD = result["TC_EXPD"];
    l_TaxCompMetaData.TC_ExpM = result["TC_ExpM"];
    l_TaxCompMetaData.TC_DateIssued = result["TC_DateIssued"];
    l_TaxCompMetaData.TC_LibraryPath = result["TC_LibraryPath"];
    l_TaxCompMetaData.TC_ScriptFunction = result["TC_ScriptFunction"];
    l_TaxCompMetaData.TC_DontAskForUpdate = result["TC_DontAskForUpdate"];

    l_Metadata["Product"] = l_TemplateMetaData;
    l_Metadata["MFMA"] = l_MFMAMetaData;
    l_Metadata["PROBE"] = l_PROBEMetaData;
    l_Metadata["NT"] = l_NTMetaData;
    l_Metadata["PFMANT"] = l_PFMANTMetaData;
    l_Metadata["TaxComp"] = l_TaxCompMetaData;
    oCW = null;
    return l_Metadata;
  }

  return {
    getMetaInfo: getMetaInfo
  };
});

var STATSTemplateInfo = (function() {

  function getFileInfoStructure() {
    return {
      FolderExists: false,
      FileExists: false,
      Compressed: false
    };
  }

  function getTemplateList() {
    var oCW = new ActiveXObject("CaseWare.Application");
    var oCWTemplateList = new Enumerator(oCW.TemplateList);

    oCWTemplateList.moveFirst();
    var oTemplates = {}; //new Array();

    while (!oCWTemplateList.atEnd()) {
      var oCWTemplate = oCWTemplateList.item();
      oTemplates[oCWTemplate.Id] = {};
      oTemplates[oCWTemplate.Id]["File"] = oCWTemplate.FilePath;
      oTemplates[oCWTemplate.Id]["Icon"] = oCWTemplate.IconPath;
      oTemplates[oCWTemplate.Id]["Name"] = oCWTemplate.Name;
      oTemplates[oCWTemplate.Id]["Type"] = oCWTemplate.Type;
      oTemplates[oCWTemplate.Id]["UsedAsDocLib"] = oCWTemplate.UsedAsDocLib;
      oTemplates[oCWTemplate.Id]["Introduction"] = oCWTemplate.Introduction;
      oTemplates[oCWTemplate.Id]["Packaged"] = oCWTemplate.Packaged;
      oTemplates[oCWTemplate.Id]["User"] = oCWTemplate.User;
      oTemplates[oCWTemplate.Id]["VersionBuild"] = oCWTemplate.VersionBuild;
      oTemplates[oCWTemplate.Id]["VersionMajor"] = oCWTemplate.VersionMajor;
      oTemplates[oCWTemplate.Id]["VersionMinor"] = oCWTemplate.VersionMinor;
      oTemplates[oCWTemplate.Id]["VersionTag"] = oCWTemplate.VersionTag;

      //getFileMetaDataStructure AND getFileInfoStructure
      var l_FileInfo = getFileInfoStructure();

      var ls_Path = oCWTemplate.FilePath;
      var ls_Name = oCWTemplate.Name;
      ls_Path = CQSGetFilePathLib(ls_Path);

      if (CQSFolderExists(ls_Path)) {
        l_FileInfo.FolderExists = true;
        var l_sPathandFileName = oCWTemplate.FilePath + ".AC";
        if (IsFileCompressed(l_sPathandFileName)) {
          l_sPathandFileName += "_";
          l_FileInfo.Compressed = true;
        }
        if (CQSFileExists(l_sPathandFileName)) {
          l_FileInfo.FileExists = true

          var l_MetaDataInfo = MetaDataInfo();
          oTemplates[oCWTemplate.Id]["metaData"] = l_MetaDataInfo.getMetaInfo(l_sPathandFileName)

        }
        oTemplates[oCWTemplate.Id]["File"] = l_FileInfo;
      }
      oCWTemplateList.moveNext();
    }
    delActiveXObj("CaseWare.Application");
    oCW = null;
    return oTemplates;
  }

  return {
    getTemplateInfo: getTemplateList
  };
});

var STATSUserSecurity = (function(CaseWareFile, CClientFileGUID, oCWClient) {

  function getUserSecurity() {
    var userSecurity = {};
    try {
      userSecurity["userName"] = oCWClient.Security.CurrentUser;
      userSecurity["userRole"] = "determine what access level the user has";
    } catch (e) {

    } finally {

    }
    return userSecurity;
  }

  return {
    getUserSecurity: getUserSecurity
  };
});

// Engagement file inforamtion
var STATSEngFileInfo = (function() {
  var HIGHESTMODE = ["None", "Builder", "Design"];
	var CURRENTPERIODTYPETYPE = ['cpttThirteen', 'cpttMonthly', 'cpttBiMonthly',
													'cpttQuarterly', 'cpttThirdly', 'cpttSemiAnnualy',
													'cpttYearly', 'cpttRandom'];
  function getFileInformation(CaseWareFile, ClientFileGUID, oCWClient) {
    var returnvalue = {};
    var workingdata = {};
    returnvalue[ClientFileGUID] = {};
    workingdata = returnvalue[ClientFileGUID];
    var l_ComputerInformation = ComputerInformation();
    var computerinfo = l_ComputerInformation.getInfo();
    var l_MetaDataInfo = MetaDataInfo();
    // Gather Level Entered
      // Entity Type
      // Year End
      // Reporting Period
      // Engagement Type
      // Entity Country
      // Framework
    var afsInformation = {};
    afsInformation["levelEntered"] = oCWClient.CaseViewData.DataGroupFormId("", "", "HIGHESTMODE") ? oCWClient.CaseViewData.DataGroupFormId("", "", "HIGHESTMODE") : 0;
		if(afsInformation["levelEntered"]==='')
			afsInformation["levelEntered"] = 0;

    afsInformation["entityType"] = oCWClient.CaseViewData.DataGroupFormId("", "", "AYENTITY").split("|")[3];
    var reportingDate = new Date(oCWClient.ClientProfile.YearEndDate);
		reportingDate = convertUTCDateToLocalDate(reportingDate);
    afsInformation["yearEnd"] = reportingDate;
		var periodtype = CURRENTPERIODTYPETYPE[oCWClient.ClientProfile.CurrentPeriodType];
    afsInformation["reportingPeriod"] = periodtype ? periodtype : 'Unknown';
    afsInformation["engagementType"] = oCWClient.CaseViewData.DataGroupFormId("", "", "AYENGAGMNT").split("|")[3];
    afsInformation["entityCountry"] = oCWClient.CaseViewData.DataGroupFormId("", "", "AYC1001").split("|")[3];
	afsInformation["piScore"] = oCWClient.CaseViewData.DataGroupFormId("", "", "PISCORE"); //GS11052016 adding the PISCORE value
    workingdata["computerInfo"] = computerinfo;
    workingdata["metaData"] = l_MetaDataInfo.getMetaInfo(CaseWareFile);
    workingdata["AFSInfo"] = afsInformation;
    var userSecurity = STATSUserSecurity(CaseWareFile, ClientFileGUID, oCWClient);
    workingdata["security"] = userSecurity.getUserSecurity();
    returnvalue.GUID = ClientFileGUID;
    return returnvalue;
  }

	function convertUTCDateToLocalDate(date) {
	    var newDate = new Date(date);
	    newDate.setMinutes(date.getMinutes() - date.getTimezoneOffset());
	    return newDate;
	}

  return {
    getFileInformation: getFileInformation
  }

});

var STATSEngFileEvent = (function(CaseWareFile, ClientFileGUID,
												oCWClient, Event, description){
  //Event: Constant
  // YEC - year End close
  // NEW - New file Created
  // UPDATE - Update Performed
  // LOCKDOWN - Lock Down performed
  function registerEventInfo() {
     var engFileInfoMain = STATSEngFileInfo();
     var returnvalue =  engFileInfoMain.getFileInformation(CaseWareFile, ClientFileGUID, oCWClient);
     returnvalue.GUID = ClientFileGUID;
     returnvalue.event = Event;
		 returnvalue.description = description ? description : '';
     return returnvalue;
  }

  return {
    registerEventInfo: registerEventInfo
  }
});

function gettemplatesinfoLib() {
	var l_ComputerInformation = ComputerInformation();
	var computerinfo = l_ComputerInformation.getInfo();
	var TemplateInfoStats = STATSTemplateInfo();
	var templateInfo = TemplateInfoStats.getTemplateInfo();
	var TemplateStatsInfo = {}
	var templatepacket = {};
	templatepacket["computerInfo"] = computerinfo;
	templatepacket["templateInfo"] = templateInfo;
	TemplateStatsInfo["TemplateStatsInfo"] = templatepacket;
	var cqsstats = new CQSStatsHTTP();
	cqsstats.registeredTemplates(TemplateStatsInfo);
}

function getclientfileinfoLib() {
// this will get client file information
	var l_ComputerInformation = ComputerInformation();
	var computerinfo = l_ComputerInformation.getInfo();
	var ClientFileGUID = CWClient.EngagementGUID.toString().toUpperCase();
	var ClientFileInfo = STATSEngFileInfo();
	var FileInfo = {};
	FileInfo["engFileInfo"] = ClientFileInfo.getFileInformation(CWClient.FileName, ClientFileGUID, CWClient);
	var cqsstats = new CQSStatsHTTP();
	//post to File information
	cqsstats.RegisterUserEngFile(FileInfo);
}

function registereventLib(Event, description) {
	var CaseWareFile = Client.FileName;
	var GUID = Client.EngagementGUID.toString().toUpperCase();
	var engEvent = STATSEngFileEvent(CaseWareFile, GUID, Client, Event, description);
	var eventSubmitInfo = {};
	eventSubmitInfo["engFileInfo"] = engEvent.registerEventInfo();
	var cqsstats = new CQSStatsHTTP();
	cqsstats.registerEvent(eventSubmitInfo);
}
