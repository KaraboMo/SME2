//////////////////////////////////////////////////////////////////////////////
// Event OnImportData (importType, importFile)
// 
// Called as first event in New From Client data. Template has been copied,
// accounting data has been imported, mapping and grouping for each account
// has not yet been validated (still contains source accounting data file
// assignments).
//
// Paramters:
//	importType - string designating the type of data being imported - currently supported values are
//				 "import.quickbooks.*" - Quickbooks file, where * is the version number (5.6,99,2001,2002,2003,2003.aus)
//				 "import.caseware"     - Existing CaseWare data
//				 "import"			   - Non-specific import
//  importFile - Full path of source file (QuickBooks .qbw or CaseWare .ac file)

function OnImportData (importType, importFile)
{
//debugger
	//get the info store doc

	var sourceClient 
	var aDBIDS
	var iYrEndFlag=0;
	//debugger;
	// Check to see that the import data source is an existing CaseWare file
	if (importType == "import.caseware") 
	{
	
		//Get hold of the metareader object
		//var oMetaReader = new ActiveXObject("MetaReader.clsMain"); this was the old method
		//new method var oMetaData = oCWApp.Clients.GetMetaData(sFilePath)
		var sCQSFFUserName = "";
		var sCQSFFPassword = "";
		var sCQSRetainISData = "";
		var sCQSRetainFormatting = "";
		var sCQSRetainAFSData = "";
		var sCQSPerformYEC = "";
		var sCQSExecuteTaskSliently = "";
		var sClientFile = importFile;
	
		
		//sCQSFFUserName = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSFFUserName");
		sCQSFFUserName = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSFFUserName");
		//sCQSFFPassword = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSFFPassword");
		sCQSFFPassword = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSFFPassword");
		//sCQSRetainISData = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSRetainISData");
		sCQSRetainISData = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSRetainISData");
		//sCQSRetainFormatting = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSRetainFormatting");
		sCQSRetainFormatting = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSRetainFormatting");
		//sCQSRetainAFSData = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSRetainAFSData");
		sCQSRetainAFSData = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSRetainAFSData");
		//sCQSPerformYEC = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSPerfomYEC");
		sCQSPerformYEC = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSPerfomYEC");
		//sCQSExecuteTaskSliently = oMetaReader.GetPropValuebyName("CWCustomProperty.CQSExecuteTaskSliently");
		sCQSExecuteTaskSliently = CQSGetMetaData(sClientFile, "CWCustomProperty.CQSExecuteTaskSliently");

		//check if file is protected
		if (Clients.Protected(importFile) == true)
		{
			
			//Check if the client is using the default username and password
			//if they are not prompt them for a password
			try
			{
				//Check if there is a username and password in the metadata
				if(sCQSExecuteTaskSliently==1)
				{
					sourceClient = Clients.Open (importFile,sCQSFFUserName,sCQSFFPassword);
				}else{
					sourceClient = Clients.Open (importFile,"SUP","SUP");
				}
			}
			catch(e)
			{
				//Call the function that will prompt the client to log into the file
				getUserNameAndPassWord()
				function getUserNameAndPassWord()
				{
					//Get the CW application object
					var application = new ActiveXObject("CaseWare.Application")
					//Get the location of CW directory
					var sApplicationPath = application.ApplicationInfo("ProgramPath")
					//Get the path of the login screen
					var sLoginHTMLDialogue = sApplicationPath+"\\Scripts\\SA IFRS\\HTML\\loginDialog.html"
					//Launch the log in screen
					var aUserNameAndPassWord = RunHTMLDialog(sLoginHTMLDialogue,"","")
					//Check if the client clicked OK or decided to cancel
					if (typeof(aUserNameAndPassWord)=="object")
					{
						//Get the username
						var sUserName = aUserNameAndPassWord[0]
						//get the password
						var sPassWord = aUserNameAndPassWord[1]
						try
						{
							//Attempt to open the file if it fails ask the client if they wish to try again
							sourceClient = Clients.Open (importFile,sUserName,sPassWord);
						}
						catch(e)
						{
							//Ask the client if theywich to try again logging into the client file
							var iResponse = MessageBox("Error", "Invalid Username or Password\nWould you like to try again?", MESSAGE_YESNO) 
							if (iResponse==7)
							{
								//If not return to the calling function
								return
							}else
							{
								//Launch the user name and password dialog again
								getUserNameAndPassWord()
							}
						}
					}else
					{
						//Client has selected to cancel this operation return to the calling function
						return
					}

				}
			}
		}
		else
		{
			sourceClient = Clients.Open (importFile);
		}
		// Open the original file
		if (sourceClient) 
		{			
			//get path for html dialog
			//var cwClient = Document.cwClient;
			//Launch html dialog
			//debugger
			if(sCQSExecuteTaskSliently==1)
			{
				var aRet = new Array();
				//Check which options need to be executed
				if(sCQSRetainISData==1)
				{
					aRet[aRet.length] = "INFOSTOREDATA|true";
				}
				
				if(sCQSRetainFormatting==1)
				{
					aRet[aRet.length] = "FIRMSETTING|true";
				}
				
				if(sCQSPerformYEC==1)
				{
					aRet[aRet.length] = "YRENDCLOSE|true";
				}
				
				if(sCQSRetainAFSData==1)
				{
					aRet[aRet.length] = "RETAININPUTDATAINAFS|true";
				}
				
			}else{
				var aRet = RunHTMLDialog(FilePath + "script\\OnFileNewDialog.html","","")
			}

			/*				oCVDataSet= sourceClient.CaseViewDataSet("FORMATTING", "*", "*")
						//enumerate the data set
						var oEnumerator = new Enumerator(oCVDataSet)
						for (;!oEnumerator.atEnd();oEnumerator.moveNext()) 
						{
							oItem = oEnumerator.item()
							//get the Group
							var sGroup = oItem.GROUP
							//get the form
							var sForm = oItem.FORM
							//get the ID
							var sID = oItem.ID
							var sData = oItem.DATA
							//write record to destination file
							MessageBox (sForm, sData, MESSAGE_YESNO)
							CaseViewData.SetGroupFormIdData(sGroup,sForm ,sID, CvDataTypeAuto, CvDataOvrNo, sData)
						}
			*/
			//get the destinatiion file
			//debugger
			//loop through the array returned by the html dialog and get the values options the client wants executed
			if(typeof(aRet)!="undefined" && aRet!=null)
			{
			var iOption = aRet.length;
			
			for (var k=0;k<iOption;k++)
			{
				var aOption = aRet[k].split("|")
				var sOptionName = aOption[0]
				var iOptionValue = aOption[1]
				if (typeof(sOptionName)!="undefined" && sOptionName!="")
				{
					//check if the client want to copy across information from the information store
					if (sOptionName=="INFOSTOREDATA" && iOptionValue=="true")
					{
						//Will need to check between the old codes and the new ones.
						sFileName = FilePath + "script\\DataBaseValues.csv" ;
						//sFileName = "C:\\Program Files\\caseware 2005\\Data\\New SA GAAP\\script\\DataBaseValues.csv";
						aDBIDS = ReadFromFile(sFileName);
						//debugger
						for (i=0;i<aDBIDS.length;i++)
						{
							if (aDBIDS[i]!="")
							{
								//split actual line items to get subsections.
								aIndivDBIDS = aDBIDS[i].split(",");
								for (p=0;p<aIndivDBIDS.length;p++)
								{
									//remove carriage return
									aIndivDBIDS[p] = aIndivDBIDS[p].replace(/(^\s*)|(\s*$)/g, "")//replace(/\r/g,"")
								}

								oldCVData = sourceClient.CaseViewData.DataGroupFormId(aIndivDBIDS[0],aIndivDBIDS[1],aIndivDBIDS[2]);
								//The line below has been commented out because if the cvdatabase value was zero, it was not entering the if statement.
								//if (typeof(oldCVData)!="undefined" && oldCVData !="")
								if (typeof(oldCVData)!="undefined" && oldCVData!=null)
								{
									CaseViewData.DeleteGroupFormId(aIndivDBIDS[0],aIndivDBIDS[1],aIndivDBIDS[2]);
									//var sDataType = sourceClient.GetGroupFormIdType
									//CM Fixed code to read data type from the CV database - above syntax was wrong
									var sDataType = sourceClient.CaseViewData.GetGroupFormIdType(aIndivDBIDS[0],aIndivDBIDS[1],aIndivDBIDS[2])
									var sData = 	 sourceClient.CaseViewData.DataGroupFormId(aIndivDBIDS[0],aIndivDBIDS[1],aIndivDBIDS[2])

									if (aIndivDBIDS[2]=="AYENTITY")
									{
										if (typeof(sData)!="undefined" && sData!="")
										{
											
											
											sData =  (sData.search("Company") != -1 ? "000001|-00001|CO|Company" : sData )

											sData =  (sData.search("Body corporate") != -1 ? "000003|-00001|BC|Body corporate" : sData )
	
											sData =  (sData.search("Corporation") != -1 ? "000002|-00001|CC|Close corporation" : sData )

											sData =  (sData.search("Business") != -1 ? "000004|-00001|SO|Sole trader" : sData )

											sData =  (sData.search("Trust") != -1 ? "000005|-00001|TR|Trust" : sData )

											sData =  (sData.search("Partnership") != -1 ? "000006|-00001|PR|Partnership" : sData )

										}
									}
									
									//Needed to save the ayengag to a different cv db value because onopening the info store the calc int he AYENGAGMNT cell was setting the AYENGAGMNT cv db value to zero
									//therfore I needed to go reset the popup for the audit report so clients have better data retention
									if (aIndivDBIDS[2]=="AYENGAGMNT")
									{
										CaseViewData.SetGroupFormIdData(aIndivDBIDS[0],aIndivDBIDS[1],"FNFEAYENGAG", sDataType, CvDataOvrNo, sData);
									}

								
									CaseViewData.SetGroupFormIdData(aIndivDBIDS[0],aIndivDBIDS[1],aIndivDBIDS[2], sDataType, CvDataOvrNo, sData);
								}
							}
						}
						
						//Excute remote script i.e. open the info store for database values to refresh
						//debugger
						//var oCWApp = Application;
						//if(oCWApp)
						//{
						//	debugger
							//var sProgramPath = oCWApp.ApplicationInfo("ProgramPath");
							/*var sFunctionName = "openAndRecalcCVDoc";
							//location of the script with the function to retain data
							var sScriptPath = sProgramPath+"Scripts\\SA IFRS\\CQS_IFRS.scp";							
							var sTempImportFile = importFile;
							var sImportFile = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("."))+"SI000000ZAFS.cvw";
							var sImportFilePath = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("\\"));
							var sTempClientName = sTempImportFile.substr((sTempImportFile.lastIndexOf("\\")+1))
							var sClientName = sTempClientName.substr(0,sTempClientName.length-3);
							var sTempTargetFilePath= FILENAME;
							var sTargetFilePath =sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("."))+"SI000000ZAFS.cvw";
							ExecuteRemoteScript(sImportFilePath, sClientName, sImportFile, sScriptPath, sFunctionName,"");*/
							
						//	var sProgramPath = oCWApp.ApplicationInfo("ProgramPath");
						//	var sFunctionName = "openAndRecalcCVDoc";
						//	var sScriptName = sProgramPath+"Scripts\\SA IFRS\\CQS_IFRS.scp";
						//	var sTempTargetFilePath= FILENAME;
						//	var sClientPath = sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("\\"));
						//	var sFileName =sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("."))+"SI000000ZAFS.cvw";
							//Get the actual CW client file being updated
							//var sClientName = sTempTargetFilePath.substr(sTempTargetFilePath.lastIndexOf("\\")+1);
							//ExecuteRemoteScript(sClientPath, sClientName, sFileName, sScriptName, sFunctionName, "");
							
							//Open the AFS
							//var sAFSFileName =sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("."))+"FSNG0000ZAFS.cvw";
							//ExecuteRemoteScript(sClientPath, sClientName, sAFSFileName, sScriptName, sFunctionName, "");
							
							
						//}
					}		

					//if the option is related to a year end close write a value to the database.
					//this value will be used later on to determine if year end close should be performed or not
					if (sOptionName=="YRENDCLOSE" && iOptionValue=="true" || iYrEndFlag==1)
					{
						//if the client has sekected that a year end close be performed set a year end close flag with a value of 1 in the CV database
						//else make it equal to 0
						CaseViewData.DataGroupFormId("FORMATTING","GLOBAL","YEC")="1";
						//script modifiction CM 17 2010 September
						//Bug this value will always be set to 0
						iYrEndFlag = 1;
					}
					else
						CaseViewData.DataGroupFormId("FORMATTING","GLOBAL","YEC")="0";
					
					//check if the client wants firmsettings copied across
					if (sOptionName=="FIRMSETTING" && iOptionValue=="true")
					{
					//debugger

						oCVDataSet= sourceClient.CaseViewDataSet("FORMATTING", "*", "*")
						//enumerate the data set
						var oEnumerator = new Enumerator(oCVDataSet)
						for (;!oEnumerator.atEnd();oEnumerator.moveNext()) 
						{
							oItem = oEnumerator.item()
							//get the Group
							var sGroup = oItem.GROUP
							//get the form
							var sForm = oItem.FORM
							//get the ID
							var sID = oItem.ID
							var sData = oItem.DATA
							//write record to destination file
							var sDataType = oItem.TYPE
							//var sDataType = sourceClient.GetGroupFormIdType(sGroup,sForm ,sID)
							
							
							//Check if the item in the source file exists in the destination file. If it does copy it to the destination file
							if(CaseViewData.ExistsGroupFormId(sGroup,sForm,sID)==1)
							{
								CaseViewData.SetGroupFormIdData(sGroup,sForm ,sID, sDataType, CvDataOvrNo, sData)
							}

						}

						//after coled pying the formatting group across, set database id COPYTEMPLATE to 1 so that formatting can be re-applied in the new file
						CaseViewData.DataGroupFormId("FORMATTING","CONTROLS","COPYTEMPLATE") = "1";
						
						//Adding a function call to a function that will retain additional cv database values
						retainAdditionalCVData(CaseViewData,sourceClient);
						
					}

					if (sOptionName=="RETAININPUTDATAINAFS" && iOptionValue=="true")
					{
						//Set the flag in the CV database to indicate that data in put cells needs to be retained
						//CaseViewData.DataGroupFormId("CQS","RETAINDATA","AFS") = "1";
						//Import data from the AFS
						
						//debugger
						//var oCvConv80Obj = new ActiveXObject("Cvconver80.CVOpen80");
						//if(oCvConv80Obj)
						//{
							//function to retain data in the AFS
							try{
								var sFunctionName = "retainAFSData";
								
								//Location of the CaseWare application
								var oCWApp = Application;
								var sProgramPath = oCWApp.ApplicationInfo("ProgramPath");
								
								//location of the script with the function to retain data
								var sScriptPath = sProgramPath+"Scripts\\SA IFRS\\CQS_IFRS.scp";							
								var sTempImportFile = importFile;
								var sImportFile = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("."))+"FSNG0000ZAFS.cvw";
								var sImportFilePath = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("\\"));
								var sTempClientName = sTempImportFile.substr((sTempImportFile.lastIndexOf("\\")+1))
								var sClientName = sTempClientName.substr(0,sTempClientName.length-3);
								var sTempTargetFilePath= FILENAME;
								var sTargetFilePath =sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("."))+"FSNG0000ZAFS.cvw";
								//oCvConv80Obj.ConvertOneCaseViewScript(sProgramPath, sImportFilePath, sClientName, sImportFile, sScriptPath, sFunctionName, sTargetFilePath, 1);
								 oCWApp.CVConvert.ConvertOneCaseViewScript(sProgramPath, sImportFilePath, sClientName, sImportFile, sScriptPath, sFunctionName, sTargetFilePath, 1);
								//oCvConv80Obj=null;
								oCWApp = null;
							}catch(e)
							{
								oCWApp = null;
							}
							//destroy the object and clear memory
							//oCvConv80Obj= null;
						//}
					}
					
					
				//debugger
					if (sOptionName=="RETAININPUTDATAINPROBE" && iOptionValue=="true")
					{
						//Set the flag in the CV database to indicate that data in put cells needs to be retained
						//CaseViewData.DataGroupFormId("CQS","RETAINDATA","AFS") = "1";
						//Import data from the AFS
						
					//debugger
						//var oCvConv80Obj = new ActiveXObject("Cvconver80.CVOpen80");
						//if(oCvConv80Obj)
						//{
							//function to retain data in the AFS
							try{
								var sFunctionName = "importData";//"importProbeDataOnFileNew";
								
								//Location of the CaseWare application
								var oCWApp = Application;
								var sProgramPath = oCWApp.ApplicationInfo("ProgramPath");
								
								//location of the script with the function to retain data
								var sScriptPath = sProgramPath+"Scripts\\SA IFRS\\CQS_IFRS.scp";							
								var sTempImportFile = importFile;
								var sImportFile = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("."))+"FIRMSET0ZAFS.cvw";
								var sImportFilePath = sTempImportFile.substr(0,sTempImportFile.lastIndexOf("\\"));
								var sTempClientName = sTempImportFile.substr((sTempImportFile.lastIndexOf("\\")+1))
								var sClientName = sTempClientName.substr(0,sTempClientName.length-3);
								var sTempTargetFilePath= FILENAME;
								//var sTargetFilePath =sTempTargetFilePath.substr(0,sTempTargetFilePath.lastIndexOf("."))+"FIRMSET0ZAFS.cvw";
								//var sTargetFilePath = sTempTargetFilePath
								var sTargetFilePath = importFile
								//debugger;
								//oCvConv80Obj.ConvertOneCaseViewScript(sProgramPath, sImportFilePath, sClientName, sImportFile, sScriptPath, sFunctionName, sTargetFilePath, 1);
								 oCWApp.CVConvert.ConvertOneCaseViewScript(sProgramPath, sImportFilePath, sClientName, sImportFile, sScriptPath, sFunctionName, sTargetFilePath, 1);
								//oCvConv80Obj=null;
								oCWApp = null;
							}catch(e)
							{
								oCWApp = null;
							}
							//destroy the object and clear memory
							//oCvConv80Obj= null;
						//}
					}
				}
			}
			}
			CaseViewData.DataGroupFormId("","","NEW_FILE_CREATED") = "1";
			//Adding a flag to show that a new file has been created.
			//this will be used to determine if the apply formatting progress bar should be launched the first time
			//the client opens their file.
			CaseViewData.DataGroupFormId("","","FORMATTINGPROGBAR") = "1";
			//close the source file
			sourceClient.Close();
		}
	}
}

function ReadFromFile(sFileName)
{
	//define variables
	var fso, oFile;
	var sGroup, sVar, sBuild, k
	var ForReading = 1;
	var aLineItems = new Array();
	var aSplitLine = new Array();
	
	sBuild = ""
	//open the file object
	fso = new ActiveXObject("Scripting.FileSystemObject");
	oFile = fso.OpenTextFile(sFileName, ForReading);
	//get file contents
	sText = oFile.ReadAll();
	//oFile.Close;
	//delActiveXObj("Scripting.FileSystemObject");

	//split on the carriage return to create an array
	aLineItems = sText.split("\n");
	//Set a count variable to create new groups.
	iVar = 0;

	return aLineItems;
}


function ExecuteRemoteScript(sClientPath, sClientName, sFileName, sScriptName, sFunctionName, sParam1)
{
//This function will execute a script via the CVConverxx.DLL
//The purpose of the function is to execute a script to a CVW document
//without prompting the user to log in.
//Parameters
//sClientPath = Path of the client file
//sClientName = Name of the client file without the path including the extension - ac or ac_
//sFileName = the caseview file that you want to execute this against, this includes the path
//                                                            as a sample "G:\Client Data\Conversion Tests\SME_2009_1[Incl Probe_2008_1 and AFS 2009_01_01]\SME_2009_1[Incl Probe_2008_1 and AFS 2009_01_01]FSNG0000ZAFS.cvw"
//ScripName = the script file name and path where the function is you want to execute, as a sample
//                                                            C:\Program Files\CaseWare\Scripts\SA IFRS\CQS_PatchLib.scp
//FunctionName = the function name that need to be executed
//sParam1 = the parameter that need to be passed in 
	//debugger
	try
	{
	//Declare the CVConverxx.dll object
	//CM - Modifying code to work with CW 2010
	//The dll is no longer supported
	//11 August 2010
	  var oCWApp = Application;
	  var sProgramPath = oCWApp.ApplicationInfo("ProgramPath"); 
	  oCWApp.CVConvert.ConvertOneCaseViewScript(sProgramPath, sClientPath, sClientName,sFileName, sScriptName,sFunctionName, sParam1, 0);
	//}
	}
	catch(e)
	{
		logError(e);
		oCWApp = null;
	}
  
}

//Retains additional CV data
function retainAdditionalCVData(CaseViewData,sourceClient)
{
	try{
	//	debugger;
		var aCVData = [["","","GROUPON"]];
								
		if(sourceClient && CaseViewData)
		{
			for(var i=0;i<aCVData.length;i++)
			{
				var sGroup = aCVData[i][0];
				var sForm = aCVData[i][1];
				var sId = aCVData[i][2];
				//Check if the item exists in the source file. If it does copy it to the destination file
				if(sourceClient.CaseViewData.ExistsGroupFormId(sGroup,sForm,sId)==1)
				{
					var sValue = sourceClient.CaseViewData.DataGroupFormId(sGroup,sForm,sId);
					CaseViewData.DataGroupFormId(sGroup,sForm,sId) = sValue;
				}
			}
			
		}
	}catch(e)
	{
		logError(e);
	}
}


function CQSGetMetaData(sFilePath, sPropertyName)
{
//debugger
//This will retrieve meta information in the file and path 
//specified, the property name is the item that will be 
//returned
	//Check if logging or debugger has been turned on
	//Removing the call to avoid stack overflow
	//checkDebugLib();
  try
  {
    var oMetaValue = "";

	try
	{
		var oMetaData = Application.Clients.GetMetaData(sFilePath);
		oMetaValue = oMetaData.item(sPropertyName).value;					  
	}
	catch(e)
	{
		try
		{
			oMetaValue = oMetaData.item("CWCustomProperty."+sPropertyName).value;
		}
		catch(e)
		{	

		}
	}

  }
  catch(e)
  {

  }
  return oMetaValue;
}




function CQSAddEditMetaData(sFilePath, sPropertyName, sValue)
{
//debugger
//This will retrieve meta information in the file and path 
//specified, the property name is the item that will be 
//returned
	//Check if logging or debugger has been turned on
	//Removing the call to avoid stack overflow
	//checkDebugLib();
  try
  { 
    var bSuccess = false;
	
	
	var sAddPropertyName = sPropertyName.split(".")[1]
	
	try
	{
		try
	   {              
		  var oMetaData = Application.Clients.GetMetaData(sFilePath);
		  if(oMetaData)
		  {
			 if(oMetaData.Exists(sPropertyName))
			 {
				oMetaData.Remove(sPropertyName)
				oMetaData.Add(sAddPropertyName,sValue)
				oMetaData.Commit()
				bSuccess = true
			 }
			 else
			 {
				oMetaData.Add(sAddPropertyName,sValue)
				oMetaData.Commit()
				bSuccess = true
			 }
			 oMetaData = null
		  }
	   }
	   catch(e)
	   {
		   if(oMetaData.Exists("CWCustomProperty."+sPropertyName))
			 {
				oMetaData.Remove("CWCustomProperty."+sPropertyName)
				oMetaData.Add(sPropertyName,sValue)
				oMetaData.Commit()
				bSuccess = true
			 }
			 else
			 {
				oMetaData.Add(sPropertyName,sValue)
				oMetaData.Commit()
				bSuccess = true
			 }
			 oMetaData = null
	   }		
			  
	}
	catch(e)
	{
		logError(e)
	}

  }
  catch(e)
  {

  }
  return bSuccess;
}
