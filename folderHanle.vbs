Dim FSO,fols,fol,curPath,fol0,MyFile,t,fPath,cPath,uPath,cFile,fols1,fol1,fils,fil,mFolder,i,j
Set FSO = CreateObject("Scripting.FileSystemObject")
SET fol = FSO.GetFolder(".")
set mFolder = fol
MsgBox fol.Path
SET fols = fol.Subfolders
'SET curPath = Environment.CurrentDirectory
'SET MyFile = Dir("C:\WINDOWS\WIN.INI")
If fols.count = 0 Then
   MsgBox "All Done", 0
   t.Close
   Set t = Nothing
   WScript.quit
End If
'MsgBox fols.count
i = 0
j = 0
for each fol0 in fols	'Foreach all subfolders of current directory
if fol0.Name <> "pic" then
    set fol = FSO.GetFolder(fol0)
    fPath = FSO.GetAbsolutePathName(".")
    'FSO.CopyFolder()
    'MsgBox fol.Path
    'MsgBox fol.ParentFolder.ParentFolder
    'MsgBox uPath
    cFile = fol.Path + "\*.mp4" 
	'MsgBox cFile
	'move subfolders out
	set fols1 = fol.Subfolders
	'i = 0
	for each fol1 in fols1
		uPath = mFolder + "\" + cstr(i) +fol1.Name
		FSO.MoveFolder fol1.Path, uPath 
		'i = i + 1
		j = j + 1
	next
	'move videos out
	set fils = fol0.files
	'i = 0
	for each fil in fils
		'MsgBox Right(fil.Name,3)
		if Right(fil.Name,3) = "mp4" or Right(fil.Name,3) = "avi" or Right(fil.Name,3) = "mkv" then
			'MsgBox fil.Name
			uPath = mFolder.Path + "\" + cstr(i) + fil.Name
			FSO.MoveFile fil.Path, uPath
			'MsgBox uPath
		'elseif Right(fil.Name,3) = "jpg" or Right(fil.Name,3) = "png" then
			'uPath = mFolder.Path + "\pic\" + cstr(i) + fil.Name
			''msgbox uPath
			'FSO.MoveFile fil.Path, uPath
		end if
		'i = i + 1
		j = j + 1
	next
	i = i + 1
    'FSO.MoveFile cFile, fol.ParentFolder
end if
Next
msgbox Cstr(j)

Sub subHandle(obfol)
	Dim fil, fil1, dc, dlm, pfil, fils,sFolders
	sFolders = obfol.Subfolders
	
End Sub

Sub folprocess(obfol)
Dim fil, fil1, dc, dlm, pfil, fils
 Set fils = obfol.files
 For Each fil in fils              '-- do files in folder
    pfil = FSO.GetAbsolutePathName(fil)
     Set fil1 = FSO.GetFile(pfil)
     dc = fil1.DateCreated
       s = InStr(1, dc, " ", 1)
       dc = left(dc, s - 1)
    dlm = fil1.datelastmodified
       s = InStr(1, dlm, " ", 1)
       dlm = left(dlm, s - 1)
         If r = dc or r = dlm Then 
            t.write pfil & VBCrLf & "created " & dc & VBCrLf & "modified " & dlm & VBCrLf & VBCrLf
        End If
    Set fil1 = Nothing
  Next
   Set fils = Nothing
End Sub