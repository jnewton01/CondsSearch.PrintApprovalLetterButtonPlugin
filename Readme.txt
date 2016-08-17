This project demonstrates creating a plugin assembly that displays all of the field
value changes to a loan when it is opened in Encompass.

To demonstrate this capability you need to fo the following:

1) Open the References for the project and, if the the EncompassObjects and 
   EncompassAutomation assembly references are broken, remove them and then
   re-add them. The assemblies will be located in your Encompass installation
   folder (e.g. C:\Program Files\Encompass).
   
2) Build the assembly to the file LoanMonitorPlugin.dll.

3) Copy only the DefaultScreenPlugin.dll file to the EncompassData\Data\Plugins
   folder on your Encompass Server (or, if your are running in Offline mode,
   the corresponding folder on your local machine).
   
4) Run Encompass and log in. Open any loan. When the loan is opened, a dialog 
   will appear with a list of the fields that are changed. Change a field value
   in the form and note that an entry is created in the monitor window.
   

   
 
