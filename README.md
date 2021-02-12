# SalesCompanion
A tool used for parsing Salesforce .html files downloaded via web browser

Run main.py from any location on a windows based machine. It will search the current user's download folder for .html files downloaded from salesforce and will parse the newest one. 
It will then wait for the user to click on the text in the Q/A field and hit ctrl+` which will automatically copy and parse the information from those fields.

Finally the parsed information will be applied to a template file (not on github) and the .html files will be destroyed. The finished lead sheet will be placed in the downloads folder.
