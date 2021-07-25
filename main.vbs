' File system object.
Dim fso
Dim wshell
dim DirFile
Set fso    = WScript.CreateObject("Scripting.FileSystemObject")
Dirfile = fso.GetParentFolderName(WScript.ScriptFullName) & "\Export\Data_Student.zip"

Set wshell = CreateObject("Shell.Application")

' Export path and file.
Dim exportPath 
Dim exportCsv
Dim exportExel

exportPath = "D:\latihan\VBS\Export\"
exportCsv  = "StudentCsv.csv"
exportExel = "StudentExel.xlsx"

' Database connection.
Dim connection

Set connection = CreateObject("ADODB.Connection")
'connection.ConnectionString = "Password=Password123;User ID=DESKTOP-T0MS1O2;Data Source=contoso_db"
connection.ConnectionString = "Driver={SQL Server};Server=DESKTOP-T0MS1O2;Database=contoso_db;Trusted_Connection=TRUE"
connection.Open()
    
' SQL to retrieve data from the person table.
Dim sqlText
sqlText = "SELECT * FROM Student"

' Retrieve the data into a record set.
Dim recordSet
Set recordSet = connection.Execute(sqlText)

' Check if the file path exists.
If fso.FolderExists(exportPath) Then

    'Delete File If The File Is Already Exist
    If fso.FileExists(exportPath & "StudentCsv.csv") Then

        fso.DeleteFile exportPath & "StudentCsv.csv"
        fso.DeleteFile exportPath & "StudentExel.xlsx"

    End If

    ' Create the CSV file.
    Dim csvFile
    Set csvFile = fso.CreateTextFile(exportPath & exportCsv, True)
    
    ' Add the header row to the CSV file.
    csvFile.WriteLine("""ID""|""LastName""|""FirstMidName""|""EnrollmentDate""")
    
    ' Data row variable.
    Dim dataRow

    ' Process the rows of data.
    Do While Not recordSet.EOF
    
        ' Construct the data row.
        dataRow = Chr(34) & recordSet("ID") & Chr(34) & "|"
        dataRow = dataRow + Chr(34) & recordSet("LastName") & Chr(34) & "|"
        dataRow = dataRow + Chr(34) & recordSet("FirstMidName") & Chr(34) & "|"
        dataRow = dataRow + Chr(34) & recordSet("EnrollmentDate") & Chr(34)
        
        ' Add the row to the CSV file.
        csvFile.WriteLine(dataRow)
        
        ' Move to the next record.
        recordSet.MoveNext
    
    Loop

    ' Close the CSV file.
    csvFile.Close
    

    ' Create the Exel file.
    Dim exelFile
    Dim Name_exel
    Name_exel = exportPath & exportExel

    Set exelFile = CreateObject("Excel.Application") 
    Set objWorkbook = exelFile.Workbooks.Add 
    exelFile.Cells(1,1).Value = "ID"
    exelFile.Cells(1,2).Value = "LastName"
    exelFile.Cells(1,3).Value = "FirstMIdName"
    exelFile.Cells(1,4).Value = "EnrollmentDate"

    For cell = 1 To 4
        exelFile.Cells(1,cell).Font.Bold = True
    Next

    Dim StudentData
    Set StudentData = connection.Execute(sqlText)
    
    Dim Countx 
    Set Countx = connection.Execute("SELECT COUNT(*)+1 FROM Student")

    For i = 2 To Countx(0)

        exelFile.Cells(i,1).Value = StudentData(0)
        exelFile.Cells(i,2).Value = StudentData(1)
        exelFile.Cells(i,3).Value = StudentData(2)
        exelFile.Cells(i,4).Value = StudentData(3)
            
        StudentData.MoveNext

    Next
    
    ' Save File
    objWorkbook.SaveAs Name_exel
    
    ' Close Sheet
    objWorkbook.Close 
    
    ' Close Exel
    exelFile.Quit
    Set exelFile = Nothing
    Set objWorkbook = Nothing

    ' Message confirming successful data export.
    WScript.Echo "Data export successful."
        
    ' Create ZIP

    ' If the name .zip is already exists delete the file and create new, if don't exists create new file .zip
    IF fso.FileExists("D:\latihan\VBS\Export\Data_Student.zip") Then

        fso.DeleteFile "D:\latihan\VBS\Export\Data_Student.zip"

    End If

    ' Create File.zip

    ' fso.CreateTextFile("D:\Data_Student.zip", True).Write "PK" & chr(5) & chr(6) & String(18, 0)

    ' Dim objshell
    ' Dim src
    ' Dim des
    ' Dim pwd

    ' Set objshell = CreateObject("Shell.Application")  
    ' Set src = objshell.NameSpace("D:\Data Student")
    ' Set des = objshell.NameSpace("D:\Data_Student.zip")

    ' des.CopyHere(src.Items)  
    
    ' Do Until des.Items.Count = src.Items.Count  
    '     Wscript.Sleep(200)  
    ' Loop  

    ' Set sh = CreateObject("WScript.Shell")
    ' sArchiveName = "sArchiveName"
    ' sWinZipLocation ="D:\latihan\VBS\Export\"
    ' sFile = "D:\latihan\VBS\Export\StudentCsv.csv"
    ' sh.Run "winzip64.exe -a -s""Password"" ""D:\latihan\VBS\Export\Data_Student.zip"" *.*", 0, True
    LocalPath = "D:\latihan\VBS\Export\"
    strWinZipDir = "C:\Program Files\WinZip\WINZIP64.exe"
    strZipFileToCreate = LocalPath & "Data_Student.zip"
    strFilesToZip  = LocalPath & "StudentCsv.csv"
    strFilesToZip1  = LocalPath & "StudentExel.xlsx"
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strWinZip = objFSO.GetFile(strWinZipDir).ShortPath
    strCommand = strWinzip & " -a -s""Password"" -r """ & strZipFileToCreate & """ " & strFilesToZip & """ " & strFilesToZip1
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(strCommand)

    'sh.Run "winzip64.exe -a -s""P4s5w07d"" ""D:\latihan\VBS\Export\Data_Student.zip"" Export\*.*", 0, True

    Wscript.echo "SuccessFully Zipped."   
       
        
    Set fso = Nothing  
    Set objshell = Nothing  
    Set src = Nothing  
    Set des = Nothing

    ' Create Email
    currentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

    emailSubject = "Lampiran Data Student"
    emailSender = "report@gmail.co.id"
    emailTo =  "Hendrik17940@gmail.com"
    emailCc = "Hendrik27.h2@gmail.com"
    emailBody = GetEmailBody
    attachment = "D:\latihan\VBS\Export\Data_Student.zip"
    sendEmail emailSubject, emailSender, emailTo, emailCc, emailBody, attachment
    
    sub sendEmail(subject, sender, recipientTo, recipientCc, bodymsg,attachment)
    wscript.echo "Sending email "
    set objEmail = CreateObject("CDO.Message")
    Set emailConfig = objEmail.Configuration
        
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") 			= 1
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory")	= "D:\latihan\VBS\Export"
    
    emailConfig.Fields.Update	
        
    objEmail.Subject = subject
    objEmail.From = sender	
    objEmail.To = recipientTo
    objEmail.Cc = recipientCc
    objEmail.HTMLBody = bodymsg
    objEmail.AddAttachment(attachment)
    
    objEmail.send
    
    wscript.echo "Email sent"
    end sub

    Function GetEmailBody()
    
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        

        GetEmailBody ="Berikut Di Lampirkan Data Student.<br /><br />"&_
                      "<strong>Automatically generated by system</strong>"
                        
    End Function


    ' Close the record set and database connection.
    recordSet.Close
    connection.Close

End If