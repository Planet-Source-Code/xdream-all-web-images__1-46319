Attribute VB_Name = "mod_frmInetScan"
Option Explicit


Public objRegExp As New RegExp

'listview
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
    ' style
    Public Const LVS_SHAREIMAGELISTS = &H40
    
    ' messages
    Public Const LVM_FIRST = &H1000
    Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
    Public Const LVM_GETITEM = (LVM_FIRST + 5)
    Public Const LVM_SETITEM = (LVM_FIRST + 6)
    Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
    Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
    Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
    
    ' LVM_GET/SETIMAGELIST wParam
    Public Const LVSIL_NORMAL = 0
    Public Const LVSIL_SMALL = 1
    
    ' LVM_GETNEXTITEM lParam
    Public Const LVNI_FOCUSED = &H1
    Public Const LVNI_SELECTED = &H2
    
    Public Type LVITEM   ' was LV_ITEM
        mask As Long
        iItem As Long
        iSubItem As Long
        state As Long
        stateMask As Long
        pszText As Long  ' if String, must be pre-allocated before filled
        cchTextMax As Long
        iImage As Long
        lParam As Long
        #If (WIN32_IE >= &H300) Then
          iIndent As Long
        #End If
    End Type
    
    ' LVITEM mask value
    Public Const LVIF_IMAGE = &H2
    
    ' LVITEM state and stateMask values
    Public Const LVIS_FOCUSED = &H1
    Public Const LVIS_SELECTED = &H2


'sub - msgbox error handle
Public Sub MsgBoxErr(errNumber As String, errDesc As String)
    On Error Resume Next
    
    If errNumber = "" Or errNumber = "0" Then
        errNumber = "<unknown>"
    End If
    
    If errDesc = "" Then
        errDesc = "<unknown>"
    End If
    
    MsgBox "Sorry, an error has occured. Please try again..." & vbCrLf & "Error " & errNumber & ": " & errDesc & ".", vbInformation + vbMsgBoxSetForeground + vbApplicationModal
    Screen.MousePointer = vbDefault
    
    If Err Then
        MsgBox "Fatal error...", vbCritical + vbMsgBoxSetForeground
    End If
End Sub


'sub - parse html for images
Public Sub subGetImages(ByRef strURL As String, ByRef strHTMLSource As String, ByRef lstLinks As ListView)
    On Error Resume Next
    
    'prepare strhtmlsource
    strHTMLSource = Replace(strHTMLSource, Chr(10), "")
    strHTMLSource = Replace(strHTMLSource, Chr(13), "")
    strHTMLSource = Replace(strHTMLSource, Chr(9), "")
    strHTMLSource = Replace(strHTMLSource, Chr(0), "")
    
    'check
    If strHTMLSource = "" Or Len(strHTMLSource) < 10 Then
        Exit Sub
    End If
    
    'vars
    Dim objMatch As Object
    Dim objMatches As Object
    Dim strIMGSource As String
    Dim strIMGPath As String
    Dim liList As ListItem
    
    'dictionary object
    Dim objDic As New Dictionary
    
    'clear list
    lstLinks.ListItems.Clear
    
    'set regexp pattern
    objRegExp.Pattern = "<IMG[^<>]+SRC=[^<>]+>"
    Set objMatches = objRegExp.Execute(strHTMLSource)
    
    'get images
    For Each objMatch In objMatches
        'prepare objmatch
        strIMGSource = CStr(objMatch)
        strIMGSource = Replace(strIMGSource, Chr(34), "")
        strIMGSource = Replace(strIMGSource, Chr(39), "")
    
        'grab source
        strIMGSource = Mid(strIMGSource, InStr(LCase(strIMGSource), "src="), Len(strIMGSource))
        objRegExp.Pattern = "src="
        strIMGSource = objRegExp.Replace(strIMGSource, "")
        objRegExp.Pattern = "[ >]{1}.*$"
        strIMGSource = objRegExp.Replace(strIMGSource, "")
        
        'determine path
        strIMGPath = fctAbsolutHtmlImgPath(strIMGSource, strURL)
        
        'determine filename
        strIMGSource = fctGetHtmlImgFilename(strIMGSource)
        
        'check vars
        If strIMGPath <> "" And strIMGSource <> "" Then
        
            'check dictionary CASE SENSITIVE
            If objDic.Exists(strIMGPath) = False Then
                'not yet in list, add
                objDic.Add strIMGPath, strIMGSource
                
                'add to list
                Set liList = lstLinks.ListItems.Add(Text:=strIMGSource)
                liList.SubItems(1) = strIMGPath
            End If
        End If
    Next
    
    'clear memory
    Set objDic = Nothing
    Set objMatch = Nothing
    Set objMatches = Nothing
    
    If Err Then
        MsgBox Err.Description
    End If
End Sub


'fct - validate url
Public Function fctValidateURL(strURL As String) As Boolean
    On Error Resume Next
    
    'check url
    If Len(strURL) < 7 Then
        fctValidateURL = False
    End If
    
    'validate url
    objRegExp.Pattern = "^http://.*\.[a-zA-Z0-9]{2,3}"
    fctValidateURL = objRegExp.Test(strURL)
    
    If Err Then
        MsgBox Err.Description
    End If
End Function


'fct - generate absolut html path for image
Public Function fctAbsolutHtmlImgPath(strIMGSource As String, strURL As String) As String
    On Error Resume Next
    
    'empty
    If Len(strIMGSource) < 5 Or strURL = "" Then
        fctAbsolutHtmlImgPath = ""
        Exit Function
    End If
    
    'vars
    Dim strRootFolder As String
    Dim strRoot As String
    
    'get rootfolder
    strRootFolder = fctGetHostName(strURL, False)
    strRoot = fctGetHostName(strURL, True)
    
CheckLink:
    'is from other source
    If fctValidateURL(strIMGSource) = True Then
        fctAbsolutHtmlImgPath = strIMGSource
        Exit Function
    End If
    
    'is path from any parentfolder
    If Left(strIMGSource, 3) = "../" Then
        strIMGSource = fctGetParentUrl(strIMGSource, True)
        strRootFolder = fctGetParentUrl(strRootFolder, False)
        GoTo CheckLink
    End If
    
    'is path from main root
    If Left(strIMGSource, 1) = "/" Then
        fctAbsolutHtmlImgPath = strRoot & Right(strIMGSource, Len(strIMGSource) - 1)
        Exit Function
    End If
    
    'is path from root folder
    If Left(strIMGSource, 1) <> "/" Then
        fctAbsolutHtmlImgPath = strRootFolder & strIMGSource
        Exit Function
    End If
    
    'path could not be retrieved
    fctAbsolutHtmlImgPath = ""
    
    If Err Then
        MsgBox Err.Description
    End If
End Function


'fct - get parent url form url
Public Function fctGetParentUrl(strURL As String, bitFile As Boolean) As String
    On Error Resume Next
    
    'vars
    Dim arrURL() As String
    Dim strOut As String
    Dim iIndex As Integer
    
    'split url
    arrURL() = Split(strURL, "/")
    
    'create parent url
    If UBound(arrURL) < 1 Then
        strOut = strURL
    Else
        If bitFile Then
            For iIndex = 1 To UBound(arrURL)
                strOut = strOut & arrURL(iIndex) & "/"
            Next iIndex
            
            'remove last "/"
            If Right(strOut, 1) = "/" Then
                strOut = Left(strOut, Len(strOut) - 1)
            End If
        Else
            For iIndex = 0 To UBound(arrURL) - 2
                strOut = strOut & arrURL(iIndex) & "/"
            Next iIndex
        End If
    End If
    
    'out
    fctGetParentUrl = strOut
    
    If Err Then
        MsgBox Err.Description
    End If
End Function


'fct - get hostname (root or folderroot)
Public Function fctGetHostName(ByVal strURL As String, bitRootOnly As Boolean)
    On Error Resume Next
    
    'vars
    Dim arrURL() As String
    Dim strOut As String
    Dim iIndex As Integer
    
    'split url
    objRegExp.Pattern = "http://"
    strURL = objRegExp.Replace(strURL, "")
    arrURL() = Split(strURL, "/")
    
    'create parent url
    If UBound(arrURL) < 1 Then
        strOut = "http://" & strURL
    Else
        If bitRootOnly Then
            strOut = "http://" & arrURL(LBound(arrURL)) & "/"
        Else
            For iIndex = 0 To UBound(arrURL) - 1
                strOut = strOut & arrURL(iIndex) & "/"
            Next iIndex
            
            'add again "http://"
            strOut = "http://" & strOut
        End If
    End If
    
    'check out
    If Len(strOut) > 0 Then
        If Right(strOut, 1) <> "/" Then
            strOut = strOut & "/"
        End If
    End If
    
    'out
    fctGetHostName = strOut

    If Err Then
        MsgBox Err.Description
    End If
End Function


'fct - get filename of source
Public Function fctGetHtmlImgFilename(strIMGSource As String)
    On Error Resume Next
    
    'empty
    If Len(strIMGSource) < 5 Then
        fctGetHtmlImgFilename = ""
        Exit Function
    End If
    
    'vars
    Dim arrSplit() As String
    arrSplit() = Split(strIMGSource, "/")
    
    'output
    fctGetHtmlImgFilename = arrSplit(UBound(arrSplit))
    
    If Err Then
        MsgBox Err.Description
    End If
End Function


'sub - download image
Public Function fctDownloadImage(strIMGSource As String, strIMGTarget As String, bitMsg As Boolean) As Boolean
    On Error Resume Next
    
    'validate source
    If fctValidateURL(strIMGSource) = False Then
        MsgBox "err, invalid source"
        Exit Function
    End If
    
    'check target
    
    'download file
    Dim bytData() As Byte
    bytData() = frmInetScan.msInet.OpenURL(strIMGSource, icByteArray)
    
    Open strIMGTarget For Binary Access Write As #1
        Put #1, , bytData()
    Close #1
    
    'check file ok
    fctDownloadImage = True
    
    If Err Then
        MsgBox Err.Description
    End If
End Function








' =============================================================================

'fct - listview - set image list
Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
    On Error Resume Next
    
    ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, iImageList, ByVal himl)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function
 

'fct - listview - get item
Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
    On Error Resume Next
        
    ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function
 

'fct - listview - set item
Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
    On Error Resume Next
        
    ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function


'fct - listview - get next item
Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
    On Error Resume Next
    
    ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags) 'MAKELPARAM(flags, 0))
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function
 
 
'fct - listview - ensure visible
Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
    On Error Resume Next
    
    ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal Abs(fPartialOK)) 'MAKELPARAM(Abs(fPartialOK), 0))
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function


'fct - listview - set item state
Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As Long, mask As Long) As Boolean
    On Error Resume Next
    
    Dim lvi As LVITEM
    lvi.state = state
    lvi.stateMask = mask
    ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function


'fct - listview - get selected item
Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
    On Error Resume Next
    
    ' Returns the index of the item that is selected and has the focus rectangle (user-defined macro)
    
    ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function

'fct - listview - set selected item
Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
    On Error Resume Next
    
    ' Selects the specified item and gives it the focus rectangle.
    ' If the listview is multiselect (not LVS_SINGLESEL), does not
    ' de-select any currently selected items (user-defined macro)
    
    ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)

    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Function


'sub - sort listview
Public Sub subSortListView(ByRef plstList As ListView, ByRef pColumnHeader As ColumnHeader)
    On Error Resume Next
    
    ' Set the ListView's SortKey to the zero-based index of the
    ' clicked column (ColumnHeader.Index is one-based).
    plstList.SortKey = pColumnHeader.Index - 1
    
    ' Toggle the column's Tag string (value), and use that value as the
    ' ListView's SortOrder (lvwAscending = 0, lvwDescending = 1)
    pColumnHeader.Tag = Abs(Not CBool(Val(pColumnHeader.Tag)))
    plstList.SortOrder = pColumnHeader.Tag
    
    ' And sort the ListView...
    plstList.Sorted = True
    plstList.SelectedItem.EnsureVisible
    
    If Err Then
        Call MsgBoxErr(Err.Number, Err.Description)
    End If
End Sub

