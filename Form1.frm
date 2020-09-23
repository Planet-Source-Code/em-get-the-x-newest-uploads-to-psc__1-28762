VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Newest PSC"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock PSCSock 
      Left            =   2640
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerPSC 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   2640
      Top             =   2040
   End
   Begin VB.TextBox TxtEntries 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   900
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "10"
      Top             =   10
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":038A
      Left            =   2520
      List            =   "Form1.frx":03A9
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0413
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LView 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   4180
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList3"
      SmallIcons      =   "ImageList"
      ColHdrIcons     =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Compatability"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Level / Author"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Views / Date Submitted"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CheckBox CheckPSCNew 
      Caption         =   "Get The          Newest Entries"
      Height          =   255
      Left            =   35
      TabIndex        =   3
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by em 2001
'http://www.em.f2s.com

'demo project showing how to communicate via the HTTP protocol
'and parse through the returned html to extract the info you want

'this demo is released into the public domain "as is" without
'warranty or guarantee of any kind. in other words, use at your own risk.
'feel free to re-use this code in your own applications, however, you
'cannot sell or re-distribute it under any circumstances.

Option Explicit

Dim PSCData As String
Dim NexPSC  As Integer
Dim PSCTot  As Integer
Dim LitM    As ListItem
Dim LVindX  As Integer

Private Sub CheckPSCNew_Click()
If CheckPSCNew.Value = 1 Then
 DoPSCSearch
End If
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 1
NexPSC = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
LView.Width = Me.Width - 115
LView.Height = Me.Height - 735 - Text1.Height - 125
Text1.Top = LView.Top + LView.Height + 25
Text1.Width = Me.Width - 115
End Sub

Private Sub LView_ItemClick(ByVal Item As MSComctlLib.ListItem)
LVindX = Item.Index
With LView.ListItems(LVindX)
  Text1 = .Tag
End With
End Sub

Private Sub LView_DblClick()
If LVindX > 0 Then
 Clipboard.Clear
 Clipboard.SetText LView.ListItems(LVindX).SubItems(1)
End If
End Sub

Private Sub PSCSock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String

PSCSock.GetData Data, vbString
PSCData = PSCData & Data
End Sub

Private Sub DoPSCSearch()
PSCData = Empty
PSCSock.Close
PSCSock.RemoteHost = "planet-source-code.com"
PSCSock.RemotePort = 80
PSCSock.Connect
TimerPSC.Enabled = True
End Sub

Private Sub PSCSock_Connect()
Dim sMSG As String
Dim firS As String
Dim laS As String

If NexPSC > 1 Then
 firS = CStr(((NexPSC - 1) * 50) + 1)
 laS = CStr(NexPSC * 50)
 If Val(laS) > Val(TxtEntries) Then laS = TxtEntries
 sMSG = "POST http://planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=" & CStr(Combo1.ListIndex) & "&?lngWId=" & CStr(Combo1.ListIndex) & "&grpCategories=-1&txtMaxNumberOfEntriesPerPage=50&optSort=DateDescending&chkThoroughSearch=&blnTopCode=False&blnNewestCode=True&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=" & firS & "&intLastRecordOnPage=" & laS & "&intMaxNumberOfEntriesPerPage=50&intLastRecordInRecordset=" & PSCTot & "&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&txtCriteria= HTTP/1.0" & vbCrLf
Else
 If Val(TxtEntries) < 50 Then
  sMSG = "GET http://planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=" & TxtEntries & "&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=" & CStr(Combo1.ListIndex) & " HTTP/1.0" & vbCrLf
 Else
  sMSG = "GET http://planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=50&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=" & CStr(Combo1.ListIndex) & " HTTP/1.0" & vbCrLf
 End If
End If
sMSG = sMSG & "Accept: */*" & vbCrLf
sMSG = sMSG & "Accept: text/html" & vbCrLf
sMSG = sMSG & "Referer: http://planet-source-code.com/vb/default.asp?lngWId=" & CStr(Combo1.ListIndex) & vbCrLf
sMSG = sMSG & "User-Agent: PSC Rocks!" & vbCrLf
sMSG = sMSG & "Host: planet-source-code.com" & vbCrLf
sMSG = sMSG & "Proxy-Connection: Close" & vbCrLf
If NexPSC > 1 Then
 sMSG = sMSG & "Content-type: application/x-www-form-urlencoded" & vbCrLf
 sMSG = sMSG & "Content-length: 22" & vbCrLf
 sMSG = sMSG & vbCrLf
 sMSG = sMSG & "cmdNextPage=++++%3E+++" & vbCrLf
End If
sMSG = sMSG & vbCrLf
PSCSock.SendData sMSG
End Sub

Private Sub PSCSock_Close()
PSCSock.Close
TimerPSC.Enabled = False
Clipboard.Clear
Clipboard.SetText PSCData
ParsePSC PSCData
End Sub

Private Sub ParsePSC(sData As String)
Dim DLTYPE As String
Dim TMPTIT As String
Dim TMPURL As String
Dim TMPDES As String
Dim COMPAT As String
Dim LEVEL As String
Dim SUBDATE As String
Dim TMPDAT As String
Dim BOOLPSC As Boolean
Dim n As Long
Dim nt As Long
Dim nt2 As Long
Dim PageTot As Integer
Dim pEntries As Integer

nt2 = InStr(1, sData, "<b>Search Results:")
If nt2 = 0 Then Exit Sub
nt = InStr(nt2, sData, "Entries")
If nt = 0 Then Exit Sub
nt2 = InStr(nt, sData, " - ")
If nt2 = 0 Then Exit Sub
pEntries = CInt(Mid$(sData, (nt + 8), nt2 - (nt + 8)))
nt = InStr(nt2, sData, " of ")
If nt = 0 Then Exit Sub
PageTot = CInt(Mid$(sData, (nt2 + 3), nt - (nt2 + 3)))
pEntries = ((PageTot - pEntries) + 1)
nt2 = InStr(nt, sData, " found")
If nt2 = 0 Then Exit Sub
PSCTot = CInt(Mid$(sData, (nt + 4), nt2 - (nt + 4)))
sData = Mid$(sData, nt2)

For n = 1 To pEntries
 nt = InStr(1, sData, "<!--descrip-->")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, "<A HREF=")
 If nt2 = 0 Then Exit Sub
 nt = InStr(nt2, sData, Chr$(62))  ' >
 If nt = 0 Then Exit Sub
 TMPURL = "http://www.planet-source-code.com" & Mid$(sData, (nt2 + 9), (nt - (nt2 + 10)))
 sData = Mid$(sData, nt)
 nt = InStr(1, sData, "src=")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, Chr$(34) & " alt=") '"
 If nt2 = 0 Then Exit Sub
 DLTYPE = Mid$(sData, (nt + 5), (nt2 - (nt + 5)))
 If DLTYPE = "/vb/scripts/images/CodeZip_small.gif" Then DLTYPE = "2"
 If DLTYPE = "/vb/images/vbicon.gif" Then DLTYPE = "3"
 If DLTYPE = "/vb/scripts/images/ArticleCopyAndPaste_small.gif" Then DLTYPE = "4"
 If DLTYPE = "/vb/scripts/images/ArticleZip_small.gif" Then DLTYPE = "4"
 If DLTYPE = "/vb/scripts/images/CIcon_small.gif" Then DLTYPE = "7"
 If DLTYPE = "" Then DLTYPE = "7"
 sData = Mid$(sData, nt2)
 nt = InStr(1, sData, Chr$(61)) '=
 If nt = 0 Then Exit Sub
 nt2 = InStr((nt + 2), sData, Chr$(34)) '"
 If nt2 = 0 Then Exit Sub
 TMPTIT = Mid$(sData, (nt + 2), (nt2 - (nt + 2)))
 sData = Mid$(sData, nt2)
 nt = InStr(1, sData, "<!--code compat-->")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, "</TD>")
 If nt2 = 0 Then Exit Sub
 COMPAT = Mid$(sData, (nt + 18), nt2 - (nt + 18))
 COMPAT = Replace(COMPAT, "&nbsp;", Chr$(32))
 sData = Mid$(sData, nt2)
 nt = InStr(1, sData, "<!--level-->")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, "</TD>")
 If nt2 = 0 Then Exit Sub
 LEVEL = Mid$(sData, (nt + 12), nt2 - (nt + 12))
 nt = InStr(1, LEVEL, "<a href=")
 If nt > 0 Then
  nt2 = InStr(nt, LEVEL, Chr$(34) & ">")
  If nt2 > 0 Then
   LEVEL = Mid$(LEVEL, 1, (nt - 1)) & Mid$(LEVEL, (nt2 + 2))
  End If
 End If
 LEVEL = Replace(LEVEL, "&nbsp;", " ")
 LEVEL = Replace(LEVEL, "<BR>", vbNullString)
 LEVEL = Replace(LEVEL, "</a>", vbNullString)
 LEVEL = Replace(LEVEL, "</", vbNullString)
 LEVEL = Replace(LEVEL, Chr$(13), vbNullString)
 LEVEL = Replace(LEVEL, Chr$(10), vbNullString)
 LEVEL = Replace(LEVEL, Chr$(9), vbNullString)
 LEVEL = Replace(LEVEL, Chr$(47), "/ ")
 sData = Mid$(sData, (nt2 + nt))
 nt = InStr(1, sData, "<TD>")
 If nt = 0 Then Exit Sub
 nt = InStr(nt, sData, "1 >")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, "</TD>")
 If nt2 = 0 Then Exit Sub
 SUBDATE = Mid$(sData, (nt + 3), nt2 - (nt + 3))
 SUBDATE = Replace(SUBDATE, "<BR>", " ")
 SUBDATE = Replace(SUBDATE, Chr$(13), vbNullString)
 SUBDATE = Replace(SUBDATE, Chr$(10), vbNullString)
 SUBDATE = Replace(SUBDATE, Chr$(9), vbNullString)
 SUBDATE = Replace(SUBDATE, Chr$(47), "/ ")
 sData = Mid$(sData, nt2)
 nt = InStr(1, sData, "<!description>")
 If nt = 0 Then Exit Sub
 nt2 = InStr(nt, sData, "&nbsp;")
 If nt2 = 0 Then Exit Sub
 nt = InStr(1, sData, "<HR>")
 If nt = 0 Then Exit Sub
 TMPDES = Mid$(sData, nt2, (nt - nt2))
 nt2 = InStr(1, TMPDES, "<a href=" & Chr$(34) & "/upload/ScreenShots")
 If nt2 > 0 Then
  TMPDES = Mid$(TMPDES, 1, (nt2 - 1))
 End If
 nt2 = InStr(1, TMPDES, "...<font size=1>")
 If nt2 > 0 Then
  TMPDES = Left$(TMPDES, (nt2 - 1))
 End If
  Debug.Print TMPDES
 TMPDES = Replace(TMPDES, Chr$(13), vbNullString)
 TMPDES = Replace(TMPDES, Chr$(10), vbNullString)
 TMPDES = Replace(TMPDES, "&nbsp;", vbNullString)
 TMPDES = Replace(TMPDES, "<br>", vbNullString)
 TMPDES = Replace(TMPDES, "<P><P>", Chr$(32))
 TMPDES = Replace(TMPDES, Chr$(9), vbNullString)
 sData = Mid$(sData, nt)
 If LView.ListItems.Count >= Val(TxtEntries) Then Exit Sub
 Set LitM = LView.ListItems.Add(, , TMPTIT, , CInt(DLTYPE))
 LitM.SubItems(1) = TMPURL
 LitM.Tag = TMPDES
 LitM.SubItems(2) = COMPAT
 LitM.SubItems(3) = LEVEL
 LitM.SubItems(4) = SUBDATE
 Set LitM = Nothing
 DoEvents
Next n

If PageTot = PSCTot Then Exit Sub

If LView.ListItems.Count >= Val(TxtEntries) Then Exit Sub

NexPSC = (NexPSC + 1)
DoPSCSearch
End Sub
