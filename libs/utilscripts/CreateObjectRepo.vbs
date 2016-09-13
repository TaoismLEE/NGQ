dim objRepo
dim strXmlFilePath
dim strTsrFilePath
dim WshShell

'Used to get the current working directory
Set WshShell = WScript.CreateObject("WScript.Shell")

'Gets path to ngq.xml and where we will create the object repo
strXmlFilePath = WshShell.CurrentDirectory & "\libs\ngq.xml"
strTsrFilePath = WshShell.CurrentDirectory & "\libs\ngq.tsr"

'The object repo object that has all its methods
set objRepo = CreateObject("Mercury.ObjectRepositoryUtil")

'Function that creates the .tsr file
objRepo.ImportFromXML strXmlFilePath, strTsrFilePath