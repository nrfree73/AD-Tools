# AD-Tools
Active Directory Tools
extension file .au3 => compiled to Autoit script exe.

3 same versions, tiny, medium and large exe (tiny no music, medium music, large exe with 2d demo)

It's a tool to automate some manipulations in Active Directory (works on DCT, CORP and domain COURRIER) or other domains but some functions will not be available such like Lbpai or Directives.. only works for domain DCT...

primary tool for DCT domain with a lot of automation or massive manipulations with files


Example: want to modify date of account massively ?
choose "Extend account" with no parameters in box "Source", choose and import a file.. => example of file: extend massively accounts with a file: 
pabc123|jj/mm/yyyy
pxyz987|jj/mm/yyyy
pzsd998|0
...

then, it will generate output file and check if account exists, and if exists get firstly previous expiration date then modify with new date..
0 means no expiration date for this account, if you put 0 ,remove expiration date...

*** Start of Users Date Extend  DCT ***  21/12/2021 09:38:00

XCBT345  |		Date d'Origine: 2021/12/31 00:00:00		|		Date prolongée:	30/06/2022  |  Resultat: OK
XSUQ972  |		Date d'Origine: 2021/12/31 00:00:00		|		Date prolongée:	30/06/2022  |  Resultat: OK
XQDI785  |		Date d'Origine: 0				              |		Date prolongée:	30/06/2022  |  Resultat: OK
XRJD190  |		Date d'Origine: 2021/12/31 00:00:00		|		Date prolongée:	30/06/2022  |  Resultat: OK
xzuy215  |		Date d'Origine: 2021/12/31 00:00:00		|		Date prolongée:	31/08/2022  |  Resultat: OK
xugj717  |		Date d'Origine: 2021/12/31 00:00:00		|		Date prolongée:	30/06/2022  |  Resultat: OK

*** End of Extend Date Users DCT ***  21/12/2021 09:38:00
