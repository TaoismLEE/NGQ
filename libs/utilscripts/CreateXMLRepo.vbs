dim objRepo
dim strXmlFilePath
dim strTsrFilePath
dim WshShell
dim fso
set fso = CreateObject("Scripting.FileSystemObject")

'Used to get the current working directory
Set WshShell = WScript.CreateObject("WScript.Shell")

'Gets path to ngq.tsr and where we will create the .xml
strXmlFilePath = WshShell.CurrentDirectory & "\libs\ngq.xml"
strTsrFilePath = WshShell.CurrentDirectory & "\libs\ngq.tsr"

'The object repo object that has all its methods
set objRepo = CreateObject("Mercury.ObjectRepositoryUtil")

'Delete old xml file before attempting xml creation
if fso.FileExists(strXmlFilePath) then
            fso.DeleteFile strXmlFilePath
        end if

'Function that creates the xml file
objRepo.ExportToXML strTsrFilePath, strXmlFilePath