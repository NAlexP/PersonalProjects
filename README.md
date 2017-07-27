# PersonalProjects
Option Explicit
Dim olMail As Outlook.MailItem
Dim s As String
Set s = olMail.Body
If olMail.Subject = "Leave Request" Then

Private objNS As Outlook.NameSpace
Private WithEvents objItems As Outlook.Items

Private Sub Application_Startup()
 
Dim objWatchFolder As Outlook.Folder
Set objNS = Application.GetNamespace("MAPI")

'Set the folder and items to watch:
' Use this for a folder in your default data file
Set objWatchFolder = objNS.GetDefaultFolder(olFolderInbox)
End Sub

Private Sub objItems_ItemAdd(ByVal Item As Object)
 
 Dim xlApp As Object
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String

 Dim strColA, strColB, strColC As String
               
' Get Excel set up
enviro = CStr(Environ("VAIO"))
'the path of the workbook
 strPath = enviro & "\Documents\Leave Requests.xlsx"
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     'Open the workbook to input the data
     Set xlWB = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWB.Sheets("Leave Requests")
    ' Process the message record
    
    On Error Resume Next
'Find the next empty line of the worksheet
rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
'needed for Exchange 2016. Remove if causing blank lines.
rCount = rCount + 1

 'collect the fields
    strColA = Item.SenderName
    
    Function FormatOutput(s)
    Dim k As Integer
    Set k = 1
    Dim re, match
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "[\d]+[.][\d]+[.][\d]+"
    re.Global = True

    For Each match In re.Execute(s)
        If IsDate(match.Value) Then
            FormatOutput = CDate(match.Value)
                        If k = 1 Then
                        strColB = match.Value
                        Set k = k + 1
                        Else
                        strColC = match.Value
                        Set k = 1
                        
                        
                        
            Exit For
        End If
    Next
    Set re = Nothing

End Function

'write them in the excel sheet
  xlSheet.Range("A" & rCount) = strColA
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("C" & rCount) = strColC
 
'Next row
  rCount = rCount + 1

     xlWB.Close 1
     If bXStarted Then
         xlApp.Quit
     End If
    
     Set xlApp = Nothing
     Set xlWB = Nothing
     Set xlSheet = Nothing
 End Function
End If
