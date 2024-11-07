<!-- #include virtual="clsUpload.asp"-->
<%

'// USE: 
'// <form method='post' name='filebrowse' id='filebrowse' encType='multipart/form-data' action='_uploadFile.asp'>
'//		<input id='inpFile1' name='inpFile1' type='File' accept='.jpg'>
'// 	<input name='inpDesc1' value='Optional text'>
'// </form>

Server.ScriptTimeout = 320 '// MAY NO LONGER BE NEEDED'

dim fileName, filePath, fileType, fileTitle

dim Upload
set Upload = new clsUpload

fileName = Upload.Fields("inpFile1").FilePath
fileTitle = Upload.Fields("inpDesc1").Value '// OPTIONAL (NOT USED IN THIS DEMO)'

if len(fileTitle) = 0 then fileTitle = fileName

select case true
	case Right(LCase(fileName), 4) = ".pdf"
		fileType = "PDF"
	case Else
		fileType = "Unknown"
end select

if fileType = "Unknown" then

	'// WRONG FILE TYPE

else

	'// UPLOAD FILE (PDF)

	filePath = server.mappath("\uploads") & "\" & fileName

	Upload("inpFile1").SaveAs filePath

end if

set Upload = nothing

%>
