# OWS-automation

### About
This is a script to automate ows web application which is used for monitoring and getting performance of network elements. The network elements include 2G technology base stations,
3G technology base stations, 4G technology base stations and Bsc (Base station controller).
This script automate escalation of fault alarms to field engineers to reduce downtime and improve availabilty. It scraps the ows web application, export base station data and escalate 
these alarms through a bulk sms platform that is also scrapped. This is achieved by repeating this action at regular interval using window's task scheduler. 
The web scrapping is carried out with selenium a python module and the data is cleaned and properly filtered by pandas module.

#### Modules
for this script to run, the following python modules should be installed;
* Pandas
* Selenium
* bs4

