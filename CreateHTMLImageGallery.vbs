'Simple html gallery maker using ImageMagic mogrify.exe
'Written by Nicklas H http://nirklars.wordpress.com

'Updated 2015-06-11

'---DECLARATIONS SECTION---

Set args = Wscript.Arguments
Set FSO = CreateObject("Scripting.FileSystemObject")
Set SHELL = CreateObject("WScript.Shell")

'Get proper current directory
SHELL.CurrentDirectory = FSO.GetParentFolderName(Wscript.ScriptFullName) 
'Declare folder to check files in
Set objFolder = FSO.GetFolder(SHELL.CurrentDirectory)
Set colFiles = objFolder.Files

'Declare global variables
Dim OutputFile 
Dim OutputFileContent
Dim FirstLineCheck

'You can change this to gallery.html or whatever you wish
OutputFile = SHELL.CurrentDirectory & "\index.htm"
'Change this if you want another thumbnail size or quality
ImageMagicArguments = "-thumbnail 200x -quality 65 -verbose"

'---PROGRAM SECTION---

'Check if a gallery already exists, if so then delete if
DelFile(OutputFile)

'Regular html gallery code stuff, this is just an example
W("<html>")
W("	<head>")
W("	<title>Image Gallery</title>")
W("		<style>  ")
W("		a:link")
W("		{")
W("			text-decoration: none;")
W("			color: black;")
W("		}")
W("		a:visited")
W("		{")
W("			color: black;")
W("		}")
W("		a:hover")
W("		{")
W("			color: red;")
W("		}")
W("		img ")
W("		{")
W("			border-style: solid;")
W("			border-width: 2px;")
W("			margin: 2px;")
W("			padding: 0px;")
W("			}")
W("		</style>")
W("	</head>")
W("	<body>")

' Breakdown of the <img> tag to be pieced together in the loop
ImagePart1 = "		<a href='"
ImagePart2 = "'><img src='thumbs/"
ImagePart3 = "'></a>"

' Go through every file in the folder and create the <img> tag and stuff
For Each objFile in colFiles
	complies = false
	
	'check file extensions
	if InStr(objFile.Name,".jpg") > 0 then 
		complies = true
	elseif InStr(objFile.Name,".png") > 0 then 
		complies = true
	elseif InStr(objFile.Name,".jpeg") > 0 then 
		complies = true
	elseif InStr(objFile.Name,".gif") > 0 then 
		complies = true
	else
		'skip
	end if
	
	if complies = true then
		W(ImagePart1 & objFile.Name & ImagePart2 & objFile.Name & ImagePart3)
	end if
	
Next

'Final part of the html code
W("	</body>")
W("</html>")

'Write the file at once to save disk access times
SaveFile()

'Check if imagemagick is there
if FSO.FileExists(SHELL.CurrentDirectory & "\mogrify.exe") then
	'Create folder and launch imagemagick
	MkDir(SHELL.CurrentDirectory & "\thumbs")
	command = Quote(SHELL.CurrentDirectory & "\mogrify.exe") & " -path " & Quote(ShortPath(SHELL.CurrentDirectory) & "\thumbs") & " " & ImageMagicArguments & " " & Quote("*")
	'start imagemagick
	SHELL.run command, 1, true
	'open the gallery in your browser
	SHELL.run Quote(OutputFile), 1, true
else
	msgbox "Unable to create thumbnail images. Please put ImageMagick mogrify.exe in the script folder and retry!"
end if

'---FUNCTIONS SECTION---

'Function to get short msdos path to work with unicode folder names
function ShortPath(myPath)
	ShortPath = FSO.GetFolder(myPath).ShortPath
end function

'Function to put quotation marks around paths
function Quote(this)
  Quote = Chr(34) & this & Chr(34)
end function

'Write a new line in the output file
Function W(strLine)
	'Skip line break on the first entry
	if FirstLineCheck = false then
		OutputFileContent = strLine
		FirstLineCheck = true
	else
		OutputFileContent = OutputFileContent & vbNewLine & strLine
	end if
End Function

'Save the output file
Function SaveFile()
    Set stream = FSO.OpenTextFile(OutputFile, 2, True)
    stream.write OutputFileContent
    stream.close
	OutputFileContent = "" 'Clear memory
End Function

'Create folder until success
function MkDir(myFolder)
	do
		'Fix env strings
		myFolder = translateEnvStr(myFolder)
		
		Err.Clear
		On Error Resume Next
		if FSO.FolderExists(myFolder) = true then 
			exit do
		else
			FSO.CreateFolder(myFolder)
			if ErrorMessage() = false then
				exit do
			end if
		end if
		WScript.Sleep 1000
	loop
end function

'Delete a file if it exists, wait if it doesnt work and retry
function DelFile(myFile)
	do
		'Fix env strings
		myFile = translateEnvStr(myFile)
		
		Err.Clear
		On Error Resume Next
		if FSO.FileExists(myFile) then
			FSO.DeleteFile(myFile)
			if ErrorMessage() = false then
				exit do
			end if
		else
			exit do
		end if
		WScript.Sleep 1000
	loop
end function

'Function that translates all EnvironmentStrings into real paths from inside a larger string. Lets call it strLargeEES.
function translateEnvStr(strLargeEES)
	'Count the number of % characters in the supplied string. This is done by removing all of the % from strLargeEES and subtract it from the original strLargeEES.
	intCharNum = Len(strLargeEES) - Len(Replace(strLargeEES, "%", ""))
	
	'Since there are two % signs for each EnvironmentString that means we divide the total number of EnvironmentStrings by...
	intExpandedStringsNum = intCharNum/2
	
	'Loop through all of the EnvironmentStrings. Because we need to translate each one separately.
	for i = 1 to intExpandedStringsNum
		'Cut out the part to the right of %
		strFirstCut = Right(strLargeEES,Len(strLargeEES)-InStr(strLargeEES,"%"))
		'Cut out the part to the left of %
		strSecondCut = Left(strFirstCut,InStr(strFirstCut,"%")-1)
		
		'The result from our cutting reveals the first EnvironmentString!
		result = "%" & strSecondCut & "%"
		'We translate this to the real folder by using a shell object
		translated = SHELL.ExpandEnvironmentStrings(result)
		
		'When we are done we replace the original EnvironmentString with the translated in strLargeEES
		strLargeEES = Replace(strLargeEES,result,translated)
	'Repeat
	next
	
	'When done return the whole translated string to the function call
	translateEnvStr = strLargeEES
end function