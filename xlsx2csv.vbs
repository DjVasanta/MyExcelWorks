Dim FileSysObj: set FileSysObj = CreateObject("Scripting.FileSystemObject")
' Current Directory of the Script
CurrentDirectory = FileSysObj.GetAbsolutePathName(".")

Set location = FileSysObj.Getlocation(CurrentDirectory)

For each file In location.Files

If FileSysObj.GetExtensionName(file) = "xlsx" Then

		pathOut = FileSysObj.BuildPath(CurrentDirectory, FileSysObj.GetBaseName(file)+".csv")
		Dim ExcelObj
		Set ExcelObj = CreateObject("Excel.Application")
		Dim WBObj
		Set WBObj = ExcelObj.Workbooks.Open(file)
		WBObj.SaveAs pathOut, 6
		WBObj.Close False
		ExcelObj.Quit
End If
Next
