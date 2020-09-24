VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInetScan 
   Caption         =   "Image Grabber"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "frmInetScan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      ToolTipText     =   "Clear formular"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Preview selected image"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "Scan url for images"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "http://www.djxdream.ch"
      ToolTipText     =   "Enter your URL to scan..."
      Top             =   240
      Width           =   3375
   End
   Begin InetCtlsObjects.Inet msInet 
      Left            =   7800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstImages 
      Height          =   3735
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   8145
   End
   Begin VB.Image imgPreview 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "frmInetScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strURL As String

'=== FRM ========================================

'frm - load
Private Sub Form_Load()
    On Error Resume Next
    
    With objRegExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With
    
    With msInet
        .RequestTimeout = 50
    End With
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'=== CMD ========================================

'cmd - scan url
Private Sub cmdScan_Click()
    'On Error Resume Next
    
    Dim strHTML As String
    
    'apply vars
    strURL = txtURL.Text
    
    'check url
    If Len(strURL) > 7 Then
        If Left(LCase(strURL), 7) <> "http://" Then
            strURL = "http://" & strURL
            txtURL.Text = strURL
        End If
    End If
    
    'validate url
    If fctValidateURL(strURL) = False Then
        MsgBox "Please enter a correct url to scan.", vbApplicationModal + vbMsgBoxSetForeground + vbInformation
        txtURL.SetFocus
        Exit Sub
    End If
        
    'init
    Screen.MousePointer = vbHourglass
    Call subEnableDisable(False)

    'status
    frmInetScan.lblStatus.Caption = "Connection to server..."

    'get html sourcecode
    strHTML = msInet.OpenURL(strURL)
    
    'status
    frmInetScan.lblStatus.Caption = "Reading images from source..."
    
    'grab images
    Call subGetImages(strURL, strHTML, lstImages)
    
    'status
    If lstImages.ListItems.Count < 1 Then
        frmInetScan.lblStatus.Caption = "Finished, no image found."
    ElseIf lstImages.ListItems.Count = 1 Then
        frmInetScan.lblStatus.Caption = "Finished, one image found."
    Else
        frmInetScan.lblStatus.Caption = "Finished, " & lstImages.ListItems.Count & " images found."
    End If
    
    'finish
    Call subEnableDisable(True)
    Screen.MousePointer = vbDefault
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'cmd - clear form
Private Sub cmdClear_Click()
    On Error Resume Next
    
    'clear items
    lstImages.ListItems.Clear
    lstImages.ToolTipText = ""
    txtURL.Text = "http://www."
    imgPreview.Picture = LoadPicture()
    lblStatus.Caption = ""
    
    'clear vars
    strURL = ""
    
    'enable objects
    Call subEnableDisable(True)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'cmd - preview image
Private Sub cmdPreview_Click()
    On Error Resume Next
    
    'init
    Screen.MousePointer = vbHourglass
    Call subEnableDisable(False)
    
    'vars
    Dim strIMGSource As String
    Dim strIMGTarget As String
    Dim intSelectedID As Integer
    
    'clear imagebox
    imgPreview.Picture = LoadPicture()
    
    'get image source from list
    intSelectedID = ListView_GetSelectedItem(lstImages.hWnd)
    If intSelectedID > -1 Then
        strIMGTarget = lstImages.ListItems.Item(intSelectedID + 1).Text
        strIMGSource = lstImages.ListItems.Item(intSelectedID + 1).SubItems(1)
        
        strIMGTarget = App.Path & "\pics\" & strIMGTarget
    Else
        strIMGTarget = ""
        strIMGSource = ""
    End If
    
    'check values
    If strIMGTarget <> "" And strIMGSource <> "" Then
    
        'status
        frmInetScan.lblStatus.Caption = "Downloading image..."
    
        'download and show picture
        If fctDownloadImage(strIMGSource, strIMGTarget, True) = True Then
            imgPreview.Picture = LoadPicture(strIMGTarget)
            imgPreview.ToolTipText = lstImages.ListItems.Item(intSelectedID + 1).Text
            
            'status
            If imgPreview.Picture = 0 Then
                frmInetScan.lblStatus.Caption = "Error, picture not found on server"
            Else
                frmInetScan.lblStatus.Caption = "Successfully downloaded image from server."
            End If
        Else
            'status
            frmInetScan.lblStatus.Caption = "Error downloading image from server."
        End If
    Else
        MsgBox "Please select an image you want to preview.", vbApplicationModal + vbMsgBoxSetForeground + vbInformation
    End If
    
    'finish
    Call subEnableDisable(True)
    Screen.MousePointer = vbDefault
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub




'=== LST ========================================

'lst - images - sorting
Private Sub lstImages_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error Resume Next
  
    Call subSortListView(lstImages, ColumnHeader)
  
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'lst - images - click
Private Sub lstImages_Click()
    On Error Resume Next
  
    Dim strIMGName As String
    Dim intSelectedID As Integer
  
    'get image source from list
    intSelectedID = ListView_GetSelectedItem(lstImages.hWnd)
    If intSelectedID > -1 Then
        strIMGName = lstImages.ListItems.Item(intSelectedID + 1).Text
    Else
        strIMGName = ""
    End If
    
    'set tooltiptext
    lstImages.ToolTipText = "'" & strIMGName & "' on " & fctGetHostName(strURL, True)
  
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'lst - images - dblClick
Private Sub lstImages_DblClick()
    On Error Resume Next
    
    Call cmdPreview_Click

    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


'=== SUB ========================================

'sub - enable disable objects
Private Sub subEnableDisable(bitEnabled As Boolean)
    On Error Resume Next
        
    'enable/disable
    txtURL.Enabled = bitEnabled
    cmdScan.Enabled = bitEnabled
    cmdClear.Enabled = bitEnabled
    
    'set properties if no items
    If lstImages.ListItems.Count < 1 Then
        cmdPreview.Enabled = False
        lstImages.Enabled = False
    Else
        cmdPreview.Enabled = bitEnabled
        lstImages.Enabled = bitEnabled
    End If
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub


