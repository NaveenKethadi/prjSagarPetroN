1. Sagar Petroleum Private Limited (ET No:110667)
=============================================
Configuration:Settings-->Configure Transactions --> External Modules
============================================================================
	On Event    : Menu Button
	Module Type : URL
	Module URL  : /prjSagarPetroN/SagarPetro/Index?companyId=$CCode
	Menu Name	: your wish
	Function Name : 

On Authorize:  
------------------------------------------------------------------------------
Configuration:Production Planning Screen -->Settings---> External Modules:
==============================================================================
2.	On Event:			On Authorize
	Module Type:		URL
	URL:	           /prjSagarPetroN/Scripts/Sagarpetro.js
	Function:			SagarpetroOnauth


   Note:Please replce DbConfig file in bin folder(bin->XMLFiles)