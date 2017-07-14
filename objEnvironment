Option Compare Database
Option Explicit

'https://blogs.technet.microsoft.com/notesfromthefield/2008/02/26/wscript-network/
'https://ss64.com/vb/network.html
'http://www.la-solutions.co.uk/content/mvba/mvba-mapped-drives-unc.htm
'http://www.tek-tips.com/viewthread.cfm?qid=1012208
'https://technet.microsoft.com/en-ca/library/cc749104(v=ws.10).aspx
'http://www.makeuseof.com/tag/see-pc-information-using-simple-excel-vba-script/
'https://sites.google.com/site/beyondexcel/project-updates/exposingsystemsecretswithvbaandwmiapi
'https://sites.google.com/site/beyondexcel/
'https://msdn.microsoft.com/en-us/library/aa390887(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/aa394585(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/aa394599(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/aa394602(v=vs.85).aspx

Private m_ProcessorsNum As Integer
Private m_LogonServer As String
Private m_OS As String
Private m_ComputerName As String
Private m_UserName As String
Private m_ComSpec As String
Private m_WinPath As String
Private m_driveList As Scripting.Dictionary

Public Property Get ComputerName() As String
    ComputerName = m_ComputerName
End Property
Public Property Get csvDriveList() As String
    csvDriveList = Join(m_driveList.Keys, ":\;") & ":\"
End Property

Public Property Get isRegisteredComputer() As String
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("select ID from Machines where" _
            & " [Computer Name] = '" & m_ComputerName & "'" _
            & " AND [Processors] = " & m_ProcessorsNum _
            & " AND [USERNAME] = '" & m_UserName & "'")

    If rst.EOF And rst.BOF Then
        isRegisteredComputer = "Unknown"
    Else
        isRegisteredComputer = rst.Fields("ID")
    End If
    rst.Close
End Property

Public Property Get driveInfo(ByVal driveLetter As String) As String
    
    With m_driveList
        If .Exists(Left(driveLetter, 1)) Then
            driveInfo = .Item(Left(driveLetter, 1))
        Else
            driveInfo = ""
        End If
    End With
End Property


Private Sub listDrives()
    Dim d As Variant
    Set m_driveList = New Scripting.Dictionary
    m_driveList.CompareMode = vbTextCompare
    
    Dim oFileSystem As New Scripting.FileSystemObject
    With oFileSystem
        For Each d In .Drives
            With d
                Select Case .DriveType
                    Case 0: m_driveList.Add Key:=.driveLetter, Item:="Unknown"
                    Case 1: m_driveList.Add Key:=.driveLetter, Item:="Removable Drive"
                    Case 2: m_driveList.Add Key:=.driveLetter, Item:="Hard Disk Drive"
                    Case 3: m_driveList.Add Key:=.driveLetter, Item:="Network Drive"
                    Case 4: m_driveList.Add Key:=.driveLetter, Item:="CDROM Drive"
                    Case 5: m_driveList.Add Key:=.driveLetter, Item:="RAM Disk Drive"
                End Select
            Debug.Print d.path
            End With
        Next
    End With
   
    
    
    Dim oWshNetwork As Object
    Set oWshNetwork = CreateObject("WScript.Network")
    Dim oDrives As Object
    Set oDrives = oWshNetwork.EnumNetworkDrives
    Dim i As Integer
    Dim driveLetter As String
    Dim uncPath As String
    
    For i = 0 To oDrives.Count - 1 Step 2
        '   Drive is oDrives.Item(i), UNC is oDrives.Item(i + 1)
            driveLetter = Left(oDrives.Item(i), 1)
            uncPath = oDrives.Item(i + 1)
            With m_driveList
            If .Exists(driveLetter) Then
                .Item(driveLetter) = .Item(driveLetter) & ":" & uncPath
            End If
            End With
    Next
   
    Set oFileSystem = Nothing
    Exit Sub

err_ParseDriveLetter:
    Select Case Err.Number
    Case 76:
        '    Path not found -- invalid drive letter or letter not mapped
        '    See VB/VBA help on topic 'Trappable Errors'
    Case Else
        MsgBox "Error no. " & CStr(Err.Number) & ": " & Err.Description & vbNewLine & _
            "Was caused by " & Err.Source, vbOKOnly Or vbExclamation, "Error in function ParseDriveLetter"
    End Select
    
End Sub
    
    
Public Sub init()
    Dim NameValuePair() As String
    Dim Indx As Integer
    
    Indx = 1
    Do
        NameValuePair = Split(VBA.Environ(Indx), "=")
        If UBound(NameValuePair) < 0 Then Exit Do
        
        Select Case NameValuePair(0)
            Case "Number_OF_PROCESSORS"
               m_ProcessorsNum = NameValuePair(1)
            Case "LOGONSERVER"
               m_LogonServer = NameValuePair(1)
            Case "Path"
               m_WinPath = NameValuePair(1)
            Case "OS"
               m_OS = NameValuePair(1)
            Case "COMPUTERNAME"
               m_ComputerName = NameValuePair(1)
            'case "ProgramFiles"
            'case "USERDOMAIN"
            Case "USERNAME"
                m_UserName = NameValuePair(1)
            Case "ComSpec"
               m_ComSpec = NameValuePair(1)
            'case "ProgramFiles"
        End Select
        Indx = Indx + 1
    Loop
    Call listDrives
    
End Sub

