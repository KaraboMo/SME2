//////////////////////////////////////////////////////////////////////////////
// Event OnCopyMapping (templateFile)
// 
// Called after a copy template of the mapping, or after a new file is created
// for a template. Called after OnImportData and before OnFileNew. Called before
// the account mappings and groupings are validated against the new template
//
// Paramters:
//	templateFile - full path of the source template

function OnCopyMapping (templateFile)
{
		//Check the entity type to define which remap file to use
		//Open the file to check the entiy
	//debugger
		if (Clients.Protected(templateFile) == true)
		{
				//Check if the client is using the default username and password
				//if they are not prompt them for a password
				try
				{
				
					sourceClient = Clients.Open (templateFile,"SUP","SUP");
				}
				catch(e)
				{
					//Call the function that will prompt the client to log into the file
					getUserNameAndPassWord(templateFile)
				}
		}
		
		else
		{
			sourceClient = Clients.Open (templateFile);
		}
		var sEntityType= "";
		sEntityType = CaseViewData.DataGroupFormId("","","AYENTITY")
		if(typeof(sEntityType)=="undefined")
		   sEntityType = "";
		   
		var sRemapFile
		var iRemap //varible defined to check if a remap file has been specified
		var iRemapSet = 0
	//search for keywords in the text and then set Remap file to use
	if (sEntityType.search(/trust/i) != "-1")
	{
		sRemapFile = ScriptPath + "remap_TR.txt";
		iRemapSet =1;

	}
	if (sEntityType.search(/corporation/i) != "-1")
	{
		sRemapFile = ScriptPath + "remap_CC.txt";
		iRemapSet =1;
	}
	if (sEntityType.search(/sole/i) != "-1")
	{
		sRemapFile = ScriptPath + "remap_SO.txt";
		iRemapSet =1;
	}
	if (sEntityType.search(/Partnership/i) != "-1")
	{
		sRemapFile = ScriptPath + "remap_SO.txt";
		iRemapSet =1;
	}
	if (iRemapSet !=1)
	{
		sRemapFile = ScriptPath + "remap.txt";
	}
	//debugger
	// used the file remap.txt to determine which numbers to remap;
	// each line has an original mapping number and a new mapping number, delimitted by a tab

	
    // build list of mapping numbers that are to be remapped
    var remap = new oRemap(sRemapFile);
//   z=0

    // go through account database remapping or removing
	/*
	for (var e = new Enumerator(Accounts);!e.atEnd();e.moveNext()) 
	{
		var account= e.item();
		var mapNo = trimall(account.Groupings.mapping);
		if (mapNo) account.Groupings.Mapping = remap.dict.Exists(mapNo) ? remap.dict.Item(mapNo) : "";
		var mapNoFlip = trimall(account.Groupings.MappingFlip);
		if (mapNoFlip) account.Groupings.MappingFlip = remap.dict.Exists(mapNoFlip) ? remap.dict.Item(mapNoFlip) : "";
	}
	*/
	iAccounts = Accounts.Count
	for (var i=1; i<= iAccounts;i++)
	{
		try
		{
			var Account =  Accounts.item(i)
			var mapNo =Account.Groupings.Mapping.replace(/(^\s*)|(\s*$)/g, "");
			var sNewMapNo =  remap.dict.Exists(mapNo) ? remap.dict.Item(mapNo) : "";
			var sType=null;
			//check if this map number is normal and only assign if it is
			if (sNewMapNo!="" ) sType = Mappings.GetIfExists(sNewMapNo).BehaviourType
				//If behavior type is not normal, change it to normal and the remap
				if (sType!=0)
				{
					Mappings.GetIfExists(sNewMapNo).BehaviourType=0;
				}
				if (sType == 0)	if (mapNo) Account.Groupings.Mapping = sNewMapNo;
			var mapNoFlip = Account.Groupings.MappingFlip.replace(/(^\s*)|(\s*$)/g, "");
			var sNewMapFlipNo =  remap.dict.Exists(mapNoFlip) ? remap.dict.Item(mapNoFlip) : "";
			var sFlipType=null;
			//check if this map number is normal and only assign if it is
			if (sNewMapFlipNo!="") sFlipType = Mappings.GetIfExists(sNewMapFlipNo).BehaviourType
			if (sFlipType == 0)	if (mapNoFlip) Account.Groupings.MappingFlip = sNewMapFlipNo

		}
		catch (e)
		{
		}
	}

	try
	{
		//Map NETINC account
		var oNETINCAcc = Accounts.Get("NETINC")
		if (oNETINCAcc)
		{
			oNETINCAcc.Groupings.Mapping = "1.5.0.340.100.000.000.800.00000.000"
		}
	}
	catch (e) {}
}

//////////////////////////////////////////////////////////////////////////////
// Object oRemap(sTextFile)
//
// Object containing dictionary of remap items
//
// Parameters:
//  sTextFile - name of text file containing mapping numbers and their remaps

function oRemap(sTextFile)
{
    try
    {
        if(!sTextFile || sTextFile.constructor != String)
            throw "Text file containing mappings and their remaps must be supplied";

        // Local Properties
        // External Properties
        this.dict = new ActiveXObject("Scripting.Dictionary");
        // Local Methods
        this.ParseLine = oRemap_ParseLine;

        // External Methods

        // init
        var ForReading = 1 ;
        var ForWriting = 2 ;
        var TristateUseDefault = -2 ;

        var fso = new ActiveXObject( "Scripting.FileSystemObject" ) ;
        var fin = fso.GetFile( sTextFile ) ;
        var ts = fin.OpenAsTextStream( ForReading, TristateUseDefault ) ;
        while( !ts.AtEndOfStream )
        {
            var oLine = this.ParseLine(ts.ReadLine()) ;
            if(!this.dict.Exists(oLine.map))
                this.dict.Add(oLine.map, oLine.remap);
        }
    }

    catch( e )
    {
        throw "Function oRemap(): " + e;
    }
}

function oRemap_ParseLine(sText)
{
	var line = new Object();

    try
    {
        if(!sText || sText.constructor != String)
            throw "Line from text file must be provided";
            
		var tabPos = sText.indexOf("\t");
        if (tabPos != -1) {
			line.map = trimall(sText.substr(0, tabPos));
			line.remap = sText.substr(tabPos + 1);
		}
		else {
			line.map = sText;
			line.remap = "";
		}
	return line;
    }
    catch( e )
    {
        throw "Function oRemap_ParseLine(): " + e;
    }
}

//////////////////////////////////////////////////////////////////////////////

function trimall( str )
{
    // remove all spaces

    var retstr = "" ;

    var pos = 0 ;
    for (var pos = 0; pos < str.length; pos++ ) {
	if (str.charAt( pos ) != " ")
		retstr += str.charAt( pos );
    }
 
    return retstr ;
}

//////////////////////////////////////////////////////////////////////////////

function getUserNameAndPassWord(templateFile)
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
			sourceClient = Clients.Open (templateFile,sUserName,sPassWord);
		}

		catch(e)
		{
			//Ask the client if theywich to try again logging into the client file
			var iResponse = MessageBox("Error", "Invalid Username or Password\nWould you like to try again?", MESSAGE_YESNO) 

			if (iResponse==7)
			{
				//If not return to the calling function
				return
			}

			else
			{
				//Launch the user name and password dialog again
				getUserNameAndPassWord()
			}
		}
	}
	
	else
	{
		//Client has selected to cancel this operation return to the calling function
		return
	}
}