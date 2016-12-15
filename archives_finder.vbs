' ARCHIVES FINDER
' archives_finder.vbs
' 
' DESCRIPTION:
' The objective of this script is to allow archivists to find groups of records
' that may be inactive because of their age.  It is designed to be run across
' large networked file systems, although it can be run across any storage device.  
' The software finds the largest possible groupings of folders that are a given number of 
' years old, based on the date last modified attribute.  
' Since this atribute can be easily and accidentally modified (e.g., someone opening
' a file and saving it), the program allows for some fuzzy math: 
' allowing a threshold where some percentage of files must be X years old.  
' The default threshold is 95%, but this can be adjusted as needed.  The default
' number of years is seven, and this can be adjusted as well, and decimal years
' (e.g., 7.5) can be used to indicate portions of a year.
' 
' Written in VB Script.  Tested on Windows XP and Windows 7.
'
' LICENSE:
' Copyright Anthony Cocciolo
' This software is licensed under the CC BY-NC-SA 3.0 license
' see: http://creativecommons.org/licenses/by-nc-sa/3.0/



Option Explicit

Dim years, zSourceDir, fso, fuzzy, earliestDate, ofile, outPaths
Dim str_size, str_age
dim default_years, default_confidence


default_years = 7
default_confidence = 95

Set FSO = CreateObject("Scripting.FileSystemObject")


' get parameters from user
years = InputBox ( "This script will locate groups of records that may be inactive, finding the largest possible groupings.  First question, how many years old should files btext (can use decimal points, e.g., 1.5 means one and a half years)?", "Enter number of years old", default_years)

if not isnumeric(years) then
	MsgBox ("Year must be numeric.  Quitting...")
	WScript.Quit
end if

if years < 0 and years > 200 then
	MsgBox ("Years must be between 0 and 200.  Quitting...")
	WScript.Quit
end if

fuzzy = InputBox ( "Second question, what percentage of files in a folder should be " & years & " years old?", "Enter percentage of years old", default_confidence)

if not isnumeric(years) then
	MsgBox ("Year must be numeric.  Quitting...")
	WScript.Quit
end if

if not isnumeric(fuzzy) then
	MsgBox ("Percentage must be numeric.  Quitting...")
	WScript.Quit
end if

if fuzzy < 1 and fuzzy > 100 then
	MsgBox ("Percentage must be between 1 and 100.  Quitting...")
	WScript.Quit
end if

fuzzy = cint (fuzzy)

zSourceDir = BrowseFolder( "", False )
' zSourceDir = "E:\whole_collection\As Retrieved"

if NOT fso.FolderExists (zSourceDir) then
	MsgBox ("Folder does not exist.  Quitting...")
	WScript.Quit
else
	MsgBox ("Press OK and the process will begin.  This may take awhile.  You will be notified when the process is complete")
end if


' compute earliest date
dim year_part, months_float
dim  g_avgdays, g_filesize, g_totalfiles
g_avgdays = 0
g_filesize = 0
g_totalfiles = 0

year_part = int (years)

months_float = years - year_part
months_float = int (months_float * 12)

earliestDate = dateadd("yyyy", year_part*-1, date)
earliestDate = dateadd("m", months_float*-1, earliestDate)

' find groups of folders
if SearchFiles (zSourceDir, g_avgdays, g_filesize, g_totalfiles) then
	if g_totalfiles > 0 then
		
		addPath csvEscape(zSourceDir) & "," & csvEscape(CStr(FormatNumber((g_avgdays / g_totalfiles) / 365, 2)))  & "," & CSVEscape(CStr(FormatNumber(g_filesize / 1024 / 1024, 2))) & "," & CSVEscape(g_totalfiles), outPaths	
	end if
end if

' output results to user
if outPaths = "" then
	MsgBox ("No paths were found matching the criteria that you specified.")
else
	dim final_output
	' Set ofile = fso.OpenTextFile ("output.csv", 2, true)
	final_output = "archives_finder.vbs ran on " & now & " to look for largest possible groups of folders that have files where " & fuzzy & "% are " & years & " years old.  The starting directory was: " & zSourceDir & "  Found paths include:" & vbNewline & "path,average years old,directory size (MB), total files" & vbNewline & outPaths
	'ofile.write outPaths
	' ofile.close
	Save2File final_output, "output.csv"

	MsgBox ("Search results complete.  Please see output.csv for a list of folders meeting your criteria.")
end if


' append path to list of paths
sub addPath (path, ByRef pathname)
	if trim(path) <> "" then
		if trim (pathname) <> "" then
			pathname = pathname & vbNewline
		end if
		pathname = pathname & path
	end if
end sub 

' recursively search for paths that have files that are old
' returns true if meets the criteria
function SearchFiles (path, ByRef days, ByRef filesize, ByRef totalFiles)
	dim totalEarlyFiles, paths, oFolder, subfolders, sf, f
	dim totalFolders, totalEarlyFolders, SearchFilesFolders, outPathsTemp
	totalFiles = 0
	totalEarlyFiles = 0
	dim days_rec, filesize_rec, totalfiles_rec
	
	days = 0
	filesize = 0
	
	Set oFolder = FSO.GetFolder(path)
    
    ' loop thru old files and count them up
    for each f in oFolder.Files
    	
    	if  (f.attributes AND 2) then
			' do nothing for hidden or system files
    	else
    		
			if (f.DateLastModified <= earliestDate) then
				totalEarlyFiles = totalEarlyFiles + 1
			end if
					
			days = days + datediff ("d", f.DateLastModified, now)	
			filesize = filesize + f.size	
			totalFiles = totalFiles + 1
						
    	end if
    next
    
    
    ' find out if the number of files meet a threshold
    SearchFiles = false
    if totalFiles > 0 then
    	if ((totalEarlyFiles / totalFiles) * 100) >= fuzzy then
    		SearchFiles = true    		
    	end if
    else if totalFiles = 0 then
    	SearchFiles = true
    end if
    end if
    
    ' search thru sub-folders for old files recursively
	SearchFilesFolders = true	
	Set subfolders = oFolder.SubFolders
   	For Each sf in subfolders		
    	if (SearchFiles (sf.Path, days_rec, filesize_rec, totalfiles_rec)) then
    		if totalfiles_rec > 0 then
    			days = days + days_rec
    			filesize = filesize + filesize_rec
    			totalFiles = totalFiles + totalfiles_rec
    			addPath csvEscape(sf.path) & "," & CSVEscape(CStr(FormatNumber((days_rec / totalFiles_rec) / 365, 2))) & "," & CSVEscape(CStr(FormatNumber(filesize_rec / 1024 / 1024, 2))) & "," & CSVEscape(totalfiles_rec), outPathsTemp 
    		end if
    	else
    		SearchFilesFolders = false
    	end if
    Next 
    
	' if all old files, return true
	' else, add paths of sub-directories that have old files
    if SearchFiles = true and SearchFilesFolders = true then
    	SearchFiles = true
    else
    	SearchFiles = false
    	addPath outPathsTemp,outPaths
    end if
end function

function csvEscape(val)
	val = Replace(val, """", """""")
	csvEscape = """" & val & """"
end function


Function BrowseFolder( myStartLocation, blnSimpleDialog )
' This function generates a Browse Folder dialog
' and returns the selected folder as a string.
'
' Arguments:
' myStartLocation   [string]  start folder for dialog, or "My Computer", or
'                             empty string to open in "Desktop\My Documents"
' blnSimpleDialog   [boolean] if False, an additional text field will be
'                             displayed where the folder can be selected
'                             by typing the fully qualified path
'
' Returns:          [string]  the fully qualified path to the selected folder
'
'
' Function written by Rob van der Woude
' http://www.robvanderwoude.com

    Const MY_COMPUTER   = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt

    ' Set the options for the dialog window
    strPrompt = "Select a folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
        Exit Function
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function

Sub Save2File (sText, sFile)
    Dim oStream
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Open
        .CharSet = "utf-8"
        .WriteText sText
        .SaveToFile sFile, 2
    End With
    Set oStream = Nothing
End Sub
