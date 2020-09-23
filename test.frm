VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "FTP"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicFiles 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.Timer TimerFTPworking 
         Interval        =   150
         Left            =   1920
         Top             =   600
      End
      Begin VB.PictureBox PicFTPlogin 
         BorderStyle     =   0  'None
         Height          =   770
         Left            =   0
         ScaleHeight     =   765
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            TabIndex        =   11
            Text            =   "bob@bob.com"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Text            =   "anonymous"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtHost 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            TabIndex        =   9
            Text            =   "ftp.microsoft.com"
            Top             =   0
            Width           =   855
         End
         Begin VB.Label FTPLabel3 
            Alignment       =   1  'Right Justify
            Caption         =   "Password"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   480
            Width           =   735
         End
         Begin VB.Label FTPLabel2 
            Alignment       =   1  'Right Justify
            Caption         =   "Username"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label FTPLabel1 
            Alignment       =   1  'Right Justify
            Caption         =   "Host"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   735
         End
      End
      Begin MSComctlLib.Toolbar ToolbarFilesFTP 
         Height          =   570
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ilFiles"
         DisabledImageList=   "ilFiles"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "GO"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarFiles 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ilFiles"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "WEB"
               Object.ToolTipText     =   "Web"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FTP"
               Object.ToolTipText     =   "ftp"
               ImageIndex      =   2
               Style           =   2
               Value           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvFTP 
         Height          =   1335
         Left            =   0
         TabIndex        =   2
         Top             =   570
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2355
         _Version        =   393217
         Indentation     =   123
         LabelEdit       =   1
         PathSeparator   =   "/"
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
      End
      Begin InetCtlsObjects.Inet icRemote 
         Left            =   1800
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1800
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0160
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":02C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0420
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0580
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":06E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0840
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilFiles 
         Left            =   1800
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":099C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":1244
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":16A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":1AF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":1F48
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":239C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":27F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":2C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":3098
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":34EC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1800
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label labFTPstat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click to Change Login"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Click to Change Login"
         Top             =   2880
         Width           =   1935
      End
   End
   Begin VB.Menu menFtpTREEfolder 
      Caption         =   "Folder"
      Visible         =   0   'False
      Begin VB.Menu menFTPgetfolder 
         Caption         =   "Refresh/Get Folder"
      End
      Begin VB.Menu menFTPUpload 
         Caption         =   "Upload to"
      End
   End
   Begin VB.Menu menFtpTREEfile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu menFTPdownload 
         Caption         =   "Download"
      End
      Begin VB.Menu menFTPdelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'syphen_2k@hotmail.com
Public strBuffer As String
Public strExecuteCmd As String
Dim strRoot As String
Dim strIndex
Dim strDeleteIndex
Dim strPUTindex
Dim strPUTname As String

Private Sub Form_Resize()
On Error Resume Next
PicFiles.Width = Me.ScaleWidth - 240
PicFiles.Height = Me.ScaleHeight - 240
End Sub

Private Sub icRemote_StateChanged(ByVal State As Integer)
If State = icResponseCompleted Then
    'retreve data from host
    strBuffer = GetFTPData(icRemote)
    If strExecuteCmd = "DIR" Then
        'fill in the lstDirItems list box with
        'the directory names just retrieved
        labFTPstat.Caption = "Getting File: finished"
        tvFTP.Nodes(1).Expanded = True
        FillDirList tvFTP
    ElseIf strExecuteCmd = "GET" Then
        labFTPstat.Caption = "Download complete"
        MsgBox "Download complete", , "FTP: download"
    ElseIf strExecuteCmd = "PUT" Then
        If tvFTP.Nodes(strPUTindex).Children <> 0 Then
            icoNum = WhatICON(strPUTname)
            tvFTP.Nodes.Add strPUTindex, tvwChild, , strPUTname, icoNum
        End If
        labFTPstat.Caption = "Upload complete"
        MsgBox "Upload complete", , "FTP: upload"
    ElseIf strExecuteCmd = "DELETE" Then
        tvFTP.Nodes.Remove (strDeleteIndex)
        labFTPstat.Caption = "Delete complete"
        MsgBox "Delete complete", , "FTP: Delete"
    ElseIf strExecuteCmd = "PWD" Then
        labFTPstat.Caption = "PWD file is " & strBuffer
        tvFTP.Nodes.Add , , , Left(strBuffer, Len(strBuffer) - 1), 1
        strIndex = 1
        strExecuteCmd = "DIR"
        icRemote.Execute , "DIR " & strBuffer
    End If
ElseIf State = 11 Then
    MsgBox "Error:", , "FTP: Error in Inet"
End If
End Sub

Function GetFTPData(icServer As Inet) As String
Dim strDataChunk As String
Dim strDataInput As String

'get a 1024-byte chunk of data for the host
strDataChunk = icServer.GetChunk(1024, icString)

Do While (Len(strDataChunk) > 0)
    'If Form1.Visible = False Then Exit Sub
    'build a buffer of received data in the
    'strDataInput variable.
    strDataInput = strDataInput & strDataChunk
    'get the next chunk
    strDataChunk = icServer.GetChunk(1024, icString)
Loop
GetFTPData = strDataInput
End Function

Sub FillDirList(lstDirList As TreeView)

Dim lonBufferPos As Long
Dim lonLastPos As Long
Dim strDirItem As String
Dim strNodeIndex As String

lonBufferPos = InStr(strBuffer, vbCrLf)
While (lonBufferPos <> 0)
    strDirItem = Mid$(strBuffer, lonLastPos + 1, _
        lonBufferPos - lonLastPos - 1)
    If strDirItem <> "" Then
        If Right(strDirItem, 1) = "/" Then
            strFile = Replace(strDirItem, "/", "")
            lstDirList.Nodes.Add strIndex, tvwChild, , strFile, 1
        Else
            lstDirList.Nodes.Add strIndex, tvwChild, , strDirItem, WhatICON(strDirItem)
        End If
    End If
    lonLastPos = lonBufferPos + 1
    lonBufferPos = InStr(lonLastPos + 1, strBuffer, vbCrLf)
Wend
If strRoot = False Then lstDirList.Nodes(strIndex).Selected = True: lstDirList.SelectedItem.Expanded = True
End Sub

Function WhatICON(fileName As String)
If Right(fileName, 5) = ".html" Or Right(fileName, 4) = ".htm" Then
    WhatICON = 3
ElseIf Right(fileName, 4) = ".jpg" Or Right(fileName, 4) = ".gif" Then
    WhatICON = 4
ElseIf Right(fileName, 4) = ".txt" Then
    WhatICON = 5
ElseIf Right(fileName, 4) = ".exe" Then
    WhatICON = 6
ElseIf Right(fileName, 4) = ".zip" Then
    WhatICON = 7
Else
    WhatICON = 2
End If
End Function

Private Sub labFTPstat_Click()
If labFTPstat.BackColor = &HFFFFFF Then
    PicFTPlogin.Visible = True
    labFTPstat.BackColor = &H808080
    labFTPstat.Caption = "click to hide login"
    PicFiles_Resize
Else
    PicFTPlogin.Visible = False
    labFTPstat.BackColor = &HFFFFFF
    tvFTP.Height = PicFiles.ScaleHeight - tvFTP.Top - labFTPstat.Height
End If
End Sub

Private Sub menFTPdelete_Click()
If MsgBox("do you realy want to delete the file " & tvFTP.SelectedItem.Text & "?", vbYesNo, "FTP: delete?") = vbYes Then
    labFTPstat.Caption = "Deleting: " & tvFTP.SelectedItem.FullPath
    strExecuteCmd = "DELETE"
    strDeleteIndex = tvFTP.SelectedItem.Index
    icRemote.Execute , "DELETE " & tvFTP.SelectedItem.FullPath
End If
End Sub

Private Sub menFTPdownload_Click()
        CommonDialog1.fileName = tvFTP.SelectedItem.Text
        CommonDialog1.ShowSave
        If CommonDialog1.fileName <> tvFTP.SelectedItem.Text Then
            labFTPstat.Caption = "Downloading: " & tvFTP.SelectedItem.FullPath
            strExecuteCmd = "GET"
            icRemote.Execute , "GET " & tvFTP.SelectedItem.FullPath & " " & CommonDialog1.fileName
        End If
End Sub

Private Sub menFTPgetfolder_Click()
'remove children
labFTPstat.Caption = "Getting File: " & tvFTP.SelectedItem.FullPath
Dim numOfChil
numOfChil = tvFTP.SelectedItem.Children
If numOfChil <> 0 Then
    If MsgBox("Do you want to refresh this folder?", vbYesNo, "FTP: refresh?") = vbNo Then Exit Sub
End If
Do While numOfChil <> 0
    tvFTP.Nodes.Remove (tvFTP.SelectedItem.Child.Index)
    numOfChil = numOfChil - 1
Loop
icRemote.Cancel
strRoot = False
strExecuteCmd = "DIR"
strIndex = tvFTP.SelectedItem.Index
icRemote.Execute , "DIR " & tvFTP.SelectedItem.FullPath & "/"
End Sub

Private Sub menFTPUpload_Click()
CommonDialog1.fileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.fileName <> "" Then
    labFTPstat.Caption = "Uploading: " & CommonDialog1.fileName
    strExecuteCmd = "PUT"
    strPUTindex = tvFTP.SelectedItem.Index
    strPUTname = CommonDialog1.FileTitle
    icRemote.Execute , "PUT " & CommonDialog1.fileName & " " & tvFTP.SelectedItem.FullPath & "/" & CommonDialog1.FileTitle
End If
End Sub


Private Sub PicFiles_Resize()
PicFTPlogin.Width = PicFiles.ScaleWidth
txtPass.Width = PicFTPlogin.ScaleWidth - txtPass.Left
txtUser.Width = txtPass.Width
txtHost.Width = PicFiles.ScaleWidth - txtHost.Left


tvFTP.Width = PicFiles.ScaleWidth
labFTPstat.Top = PicFiles.ScaleHeight - labFTPstat.Height
labFTPstat.Width = PicFiles.ScaleWidth - labFTPstat.Left
ToolbarFilesFTP.Left = PicFiles.ScaleWidth - ToolbarFilesFTP.Width
PicFTPlogin.Top = labFTPstat.Top - PicFTPlogin.Height
If PicFTPlogin.Visible = False Then
    tvFTP.Height = PicFiles.ScaleHeight - tvFTP.Top - labFTPstat.Height
Else
    tvFTP.Height = PicFTPlogin.Top - tvFTP.Top
End If
End Sub

Private Sub TimerFTPworking_Timer()
If icRemote.StillExecuting = True Then
    If ToolbarFilesFTP.Buttons(2).Image = 11 Then
        ToolbarFilesFTP.Buttons(2).Image = 4
    Else
        ToolbarFilesFTP.Buttons(2).Image = ToolbarFilesFTP.Buttons(2).Image + 1
    End If
Else
    If ToolbarFilesFTP.Buttons(2).Image <> 8 Then ToolbarFilesFTP.Buttons(2).Image = 8
End If
End Sub
Private Sub ToolbarFiles_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "WEB" Then
    'hide ftp stuf
    tvFTP.Visible = False
    ToolbarFilesFTP.Visible = False
    labFTPstat.Visible = False
    labFTPstat.BackColor = &HFFFFFF
    PicFTPlogin.Visible = False
ElseIf Button.Key = "FTP" Then
    tvFTP.Visible = True
    ToolbarFilesFTP.Visible = True
    labFTPstat.Visible = True
    labFTPstat.BackColor = &HFFFFFF
    PicFTPlogin.Visible = False
End If
End Sub

Private Sub ToolbarFilesFTP_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "GO" Then
    With icRemote
    .AccessType = icDirect
    .RequestTimeout = 60
    .Protocol = icFTP
    .URL = txtHost.Text
    .UserName = txtUser.Text
    .Password = txtPass.Text
    tvFTP.Nodes.Clear
    strRoot = True
    strExecuteCmd = "PWD"
    labFTPstat.Caption = "Geting Defult file name"
    .Execute , "PWD"
    End With
ElseIf Button.Key = "Stop" Then
    If icRemote.StillExecuting = True Then icRemote.Cancel
End If
End Sub

Private Sub tvFTP_DblClick()
On Error GoTo exit1
If icRemote.StillExecuting = True Then MsgBox "Sorry can not open that file/folder im still working on your last request", , "FTP: error": Exit Sub
If tvFTP.SelectedItem.Image = 1 And tvFTP.SelectedItem.Index <> 1 Then
menFTPgetfolder_Click
ElseIf tvFTP.SelectedItem.Index <> 1 Then
    menFTPdownload_Click
End If
exit1:
Exit Sub
End Sub

Private Sub tvFTP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo exit1
If Button = 2 Then
    If tvFTP.SelectedItem.Image = 1 Then
        PopupMenu menFtpTREEfolder
    Else
        PopupMenu menFtpTREEfile
    End If
End If
exit1:
Exit Sub
End Sub
