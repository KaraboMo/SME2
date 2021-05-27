//////////////////////////////////////////////////////////////////////////////
// Event OnFileNew ()
// 
// Called after the Create New File or Create New File from Client Data is 
// completed. Called after initial creation of the layout windows, but before
// the client profile is launched (in the case of Create New File).

function OnFileNew()
{

	//RunHTMLDialog(sHTMLFileName As String, argument, [Flags])

	/*
		 //Get the CopyTemplate object
	var copyTemplateObj = Application.CopyTemplate;	
	
		//Get the source file and set which documents should be copied	
	var oSource = Application.Clients.Open(oSourcePath);	
		//Get the documents collection
	var documentsCollection = oSource.Documents;
		// Reset all documents to not be copied
	documentsCollection.SelectAllCopyFlag(0);	
		// Get the document with id "1" and set it to be copied for a copy template	
			var aDocuments("SI000000ZAFS", "FSNG0000ZAFS")
			for (var i=0 ;i <=aDocuments.length ;i++ )
			{	var document = documentsCollection.Get(aDocuments[i]);
				if (document != null)
				{
					document.CopyFlag = 1;	
						// Close the source file
					Application.Clients.Close(oSourcePath);

						// Set the balance to also be copied
					copyTemplateObj.CopyTrialBalance = true;

						// Set the source file for the copy template
					copyTemplateObj.CompleteSourceFileName = oSourcePath;	
						// Set the destination directory for the copy template
					copyTemplateObj.CompleteDestinationFileName = Client.FileName;

						//Perform the CopyTemplate
					try
					{
						copyTemplateObj.DoCopy();
					}

					catch (exception)
					{
					}    	
				}
				else
				{
					MessageBox ("Error", "Document with id " + aDocuments[i] +" does not exist", 0);
				}

				
			}
	*/
	//Mo - 14/11/2014 - I need to set these 2 1 so that when applyformatiing kicks in when the user opens the file after a "file new" then the content mngt ribbon only shows
	CaseViewData.DataGroupFormId("FORMATTING","PRINTING","BALCHK") = "0:1"
	CaseViewData.DataGroupFormId("FORMATTING","PRINTING","COLUMNSET") = "0:1"
	
	var e = new Enumerator(Accounts);
	if (!e.atEnd()) { // Accounts exist, so ask if year end close should be perfomed
		//if (MessageBox ("Year End Close", "Do you wish to perform a year end close on the existing data?", MESSAGE_YESNO) == MESSAGE_RESULT_YES) {
		//06/12/2006 - CM - We no longer need the messagebox above this process is controlled by a database id
		//set by the client during the file new process
		if (CaseViewData.DataGroupFormId("FORMATTING","GLOBAL","YEC")==1)
		{
			var yec = YearEndClose;
			yec.DestinationFile = Client;
			yec.UpdatePriorYearBalanceData = true;
			yec.UpdateCaseViewRollForwardCells = true;
			//currently code below does not work. this means our files do not auto roll data into the next year.
			//yec.UpdateNextYearOpeningBalanceData = true;
			//yec.BalanceType= 1;
			yec.DoYearEndClose();
		}
	}
}