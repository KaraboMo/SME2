


//debugger
//reset the completed tickboxes on the getting started window by resetting the ffg flag
if(CaseViewData.ExistsGroupFormId("","","CWAWFPROG"))
{
	CaseViewData.DataGroupFormId("","","CWAWFPROG") = '{"done":{"welcome":{"value":0},"filesettings":{"value":0},"importing":{"value":0},"mapping":{"value":0},"statement":{"value":0},"process":{"value":0},"prepare":{"value":0},"yeclose":{"value":0}}}'
}

//on rollforward when you do a file new from existing you also have to refresh the getting started window
setGettingStartedLayout()

function setGettingStartedLayout()
{
//debugger
//only run on Africa template
//only set the Africa Process Flow Layout once
//code extended:
//set the work flow layout for South African files that are of type FULL_IFRS, SME, and SAIPA
//only set the Africa Process Flow Layout once
try
  { 
  
		var bRun = false
		
		var sCountryPropertyName = "CWCustomProperty.Country"	
		var sProductPropertyName = "CWCustomProperty.Product"
		var sLayoutPropertyName = "CWCustomProperty.GettingStarted"
		
		var sCountry = ""
		var sProduct = ""
		var sLayoutSet = ""
		var sGettingStartedLayoutVersion = 2
		
	var sFilePath = Client.FileName
	
	//var oMetaData = cwClients.GetMetaData("Samp01.ac");
	var oMetaData = Clients.GetMetaData(sFilePath);
	//enumerate the data set
	var oEnumerator = new Enumerator(oMetaData)
	for (;!oEnumerator.atEnd();oEnumerator.moveNext()) 
	{
		oItem = oEnumerator.item()
		var sName = oItem.Name
		var sValue = oItem.value
		
		if(sName==sCountryPropertyName)
			sCountry = sValue
			
		if(sName==sProductPropertyName)
			sProduct = sValue
			
		if(sName==sLayoutPropertyName)
			sLayoutSet = sValue
			
	}
  
  
		//set the work flow layout for Africa files
		if(sCountry == "AFRICA")
			bRun = true
		
		//set the work flow layout for South African files that are of type FULL_IFRS, SME, and SAIPA
		if(sCountry == "ZA")
		{
			if(sProduct == "FULL_IFRS" || sProduct == "SME" || sProduct == "SAIPA")
				bRun = true
		}
		
		//cater for scenarios where existing SME and SAIPA files in the market have a blank country value
		if(sCountry == " ")
		{
			if(sProduct == "FULL_IFRS" || sProduct == "SME" || sProduct == "SAIPA")
				bRun = true
		}
		
		
		if(bRun)
		{
			/*if(sLayoutSet == sGettingStartedLayoutVersion)
			{
				return
			}
			else
			{*/

				/*
				1. add automap button record (to view manually, must be in template file and select Tools->template toolbar
				*/

				//var oClient = document.cwClient
				var oToolBarButtonsCollection = Client.TemplateToolbarButtons
				
				var sScriptString = 'if(Client.ClientOptions.Layout!="GETTINGSTARTED") \n{\n\tClient.ClientOptions.Layout="GETTINGSTARTED"; \n}else{ \n\tClient.ClientOptions.Layout="0"; \n}'
				
				var oToolBarButton = oToolBarButtonsCollection.Get("LO")
				if(oToolBarButton)
					oToolBarButtonsCollection.Remove("LO")
				
				//add tutorial button record
				var oToolBarButton = oToolBarButtonsCollection.Add("LO")
				if(oToolBarButton)
				{
					oToolBarButton.Text = "Getting Started"
					oToolBarButton.ExDescription = 'The "Getting Started" guides you through the process of preparing your engagement file.'
					oToolBarButton.Script = sScriptString
				}

				
				/*
				2. view->toolbars->template toolbar. set this to ticked by script so that the automap button shows in the toolbar
				*/	
		
				var tmp_button = oToolBarButtonsCollection.Add("TMP");
				tmp_button.Url = "cw:templatetoolbar?visible=show";
				tmp_button.Execute();
				oToolBarButtonsCollection.Remove("TMP");
		
		
				/*
				3. run the script code in the script section of the automap button record (set the “CaseWare Africa Process Flow” Layout)
				*/
				
				
				Client.ClientOptions.Layout="GettingStarted";
				
				//oMetaData.Add("GettingStarted",sGettingStartedLayoutVersion)
				//oMetaData.Commit()

				
			//}
		}


  }
  catch(e)
  {
    //logError(e);
  }
  
}
