# PACModeSLogging
Auto log into GFIA database from PlanePlotter

Application runs alongside PlanePlotter using VB.Net code to retrieve alist of aircraft from the Aircraft table in basestation.sqb which have a UserTag string containing RQ or Ps 

This adds to a dropdown list which the user can click on any of the items to automatically log (after a confirmation prompt) into the logllp table in privatelogs.mdb.  For military aircraft additional data is logged into loglls table in privatelogs.mdb

The user has a configuration option which specifies the folder location of the basestation.sqb database, whether or not they log on the basis of User Tag or the Interested flag. They also specify the location they wish to use as the place where they logged the aircraft details, for example Home or Heathrow. There is the option to set the refresh rate i.e. how often the application checks for new aircraft to add to the dropdown list.  The final option is whther the application panel should always remain on top of all other windows.

Once running the user can disable sounds which play each time a new RQ or Ps tageed aircraft is picked up by PlanePlotter. They can also clear the drop down list at any time.



