ve 
         Caption         =   "Save"
      End
      Begin VB.Menu Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeys 
         Caption         =   "Keys"
      End
      Begin VB.Menu mnuOpenKeysData 
         Caption         =   "Open Key Data"
      End
      Begin VB.Menu mnuSaveKeyData 
         Caption         =   "Save Key Data"
      End
      Begin VB.Menu Sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaintSnapShot 
         Caption         =   "Paint  SnapShot"
      End
      Begin VB.Menu mnuMovie 
         Caption         =   "Movie"
      End
      Begin VB.Menu mnuMovieTime 
         Caption         =   "MovieTime 05 Sec"
         Begin VB.Menu mnuMTime 
            Caption         =   "03 Sec"
            Index           =   0
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "04 Sec"
            Index           =   1
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "05 Sec"
            Index           =   2
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "06 Sec"
            Index           =   3
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "07 Sec"
            Index           =   4
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "08 Sec"
            Index           =   5
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "09 Sec"
            Index           =   6
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "10 Sec"
            Index           =   7
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "15 Sec"
            Index           =   8
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "20 Sec"
            Index           =   9
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "25 Sec"
            Index           =   10
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "30 Sec"
            Index           =   11
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "45 Sec"
            Index           =   12
         End
         Begin VB.Menu mnuMTime 
            Caption         =   "60 Sec"
            Index           =   13
         End
      End
      Begin VB.Menu mnuOpenSnapShot 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSaveSnapShot 
         Caption         =   "Save"
      End
      Begin VB.Menu Sep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ResizeTextBox()
 txtKeys.Width = Me.Width - 300
 txtKeys.Height = Me.Height / 2
 txtKeys.Visible = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then PopupMenu mnuFile
End Sub

Private Sub mnuClear_Click()
 txtKeys.Visible = False
 Me.Picture = LoadPicture()
End Sub

Private Sub mnuExit_Click()
 tmrMovie.Enabled = False
 Unload Me
End Sub

Private Sub mnuKeys_Click()
 Dim objF As New FileSystemObject
 Dim objT As TextStream
 On Error GoTo EH
 
 Set objT = objF.OpenTextFile(App.Path & "\WinsKBoard.spy", ForReading)
 txtKeys.Text = objT.ReadAll
 ResizeTextBox
 DoEvents
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description
 
End Sub

Private Sub mnuMovie_Click()
On Error GoTo EH

 If Query.cboClientIP = Query.cboRemoteIP Then
   MsgBox "Change your Remote IP" & vbCrLf & "It will not scan Local Screen"
   Exit Sub
 End If
 
 tmrMovie.Enabled = Not tmrMovie.Enabled
 mnuMovie.Checked = Not mnuMovie.Checked
 
 Exit Sub
EH:
 MsgBox Err.Description & vbCrLf & Err.Number
End Sub

Private Sub mnuMTime_Click(Index As Integer)
 'Dim intC As Integer
 tmrMovie.Interval = Val(mnuMTime(Index).Caption) * 1000
 'For intC = 0 To mnuMTime.Count - 1
  ' mnuMTime(intC).Checked = False
 'Next
 'DoEvents
 'mnuMTime(Index).Checked = True
 mnuMovieTime.Caption = "MovieTime " & mnuMTime(Index).Caption
 
End Sub

Private Sub mnuOpen_Click()
 Dim strFName As String
 On Error GoTo EH
  txtKeys.Visible = False
  cdlPicture.FilterIndex = 3
  cdlPicture.ShowOpen
  strFName = cdlPicture.FileName
  DoEvents
  Paint strFName
  Exit Sub
  
EH:
  MsgBox Err.Number & Err.Description
End Sub

Private Sub Paint(strFile As String)
 Dim X1 As Integer, X2 As Integer, xx As Integer
 Dim Y1 As Integer, Y2 As Integer, yy As Integer
 
 Dim objFile As New FileSystemObject
 Dim objTextStream As TextStream
 
 Dim lngColorCounter As Long, lngColorValue As Long
 Dim C() As String
 
 On Error GoTo EH
 If Trim(strFile) = vbNullString Then Exit Sub
 Set objTextStream = objFile.OpenTextFile(strFile, ForReading)
 
 
 With objTextStream
   X1 = CInt(.ReadLine): Y1 = CInt(.ReadLine)
   X2 = CInt(.ReadLine): Y2 = CInt(.ReadLine)
   
   For xx = X1 To X2
     For yy = Y1 To Y2
     
        If lngColorCounter = 0 Then
            C = Split(.ReadLine, ".")
            lngColorCounter = CLng(C(0))
            lngColorValue = CLng(C(1))
        End If

       SetPixelV Me.hdc, xx, yy, lngColorValue
       lngColorCounter = lngColorCounter - 1
   Next yy, xx
   
 End With
 
 Exit Sub
EH:
   MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub mnuOpenSnapShot_Click()
 Dim strFName As String
 On Error GoTo EH
  txtKeys.Visible = False
  cdlPicture.FilterIndex = 4
  cdlPicture.ShowOpen
  strFName = cdlPicture.FileName
  DoEvents
  Me.Picture = LoadPicture(strFName)
  Exit Sub
  
EH:
  MsgBox Err.Number & Err.Description
End Sub

Private Sub mnuOpenKeysData_Click()
  Dim strFName As String
  Dim objF As New FileSystemObject
  Dim objT As TextStream
  
 On Error GoTo EH
  cdlPicture.FilterIndex = 2
  cdlPicture.ShowOpen
  strFName = cdlPicture.FileName
  DoEvents
  If strFName = vbNullString Then Exit Sub
  
  Set objT = objF.OpenTextFile(strFName)
  txtKeys.Text = objT.ReadAll
  ResizeTextBox
  DoEvents
  Exit Sub
  
EH:
  MsgBox Err.Number & Err.Description

End Sub

Private Sub mnuPaint_Click()
 txtKeys.Visible = False
 DoEvents
 Paint App.Path & "\" & "Map.spy"
End Sub

Private Sub mnuPaintSnapShot_Click()
 On Error GoTo EH
 txtKeys.Visible = False
 Me.Picture = LoadPicture(App.Path & "\WinMap.bmp")
 Exit Sub
EH:
 MsgBox Err.Description & vbCrLf & Err.Number
End Sub

Private Sub mnuSave_Click()
  Dim objF As New FileSystemObject
  Dim strDest As String
  On Error GoTo EH
   txtKeys.Visible = False
   cdlPicture.FilterIndex = 3
   cdlPicture.ShowSave
   strDest = cdlPicture.FileName
      If Trim(strDest) = vbNullString Then Exit Sub
   objF.CopyFile App.Path & "\" & "Map.spy", strDest
  Exit Sub
  
EH:
  MsgBox Err.Number & Err.Description
End Sub

Private Sub mnuSaveSnapShot_Click()
  On Error GoTo EH:
  cdlPicture.FilterIndex = 4
  cdlPicture.ShowSave
  SavePicture Me.Picture, cdlPicture.FileName
  
  Exit Sub
EH:
  Select Case Err.Number
   Case 380: MsgBox "There is no image at your Window! " & vbCrLf & " Click Paint SnapShot and then Save the file! "
   Case Else:  MsgBox Err.Description & vbCrLf & Err.Number
  End Select
  
End Sub

Private Sub mnuSaveKeyData_Click()
 Dim objF As New FileSystemObject
  Dim strDest As String
  On Error GoTo EH
   txtKeys.Visible = False
   cdlPicture.FilterIndex = 2
   cdlPicture.ShowSave
   strDest = cdlPicture.FileName
      If Trim(strDest) = vbNullString Then Exit Sub
   objF.CopyFile App.Path & "\" & "WinsKBoard.spy", strDest
  Exit Sub
  
EH:
  MsgBox Err.Number & Err.Description
End Sub

Private Sub tmrMovie_Timer()
 Query.FastCapture (vbLeftButton)
 DoEvents
 PlaySound "Camera.wav"
 Query.FastCapture (vbRightButton)
 DoEvents
 mnuPaintSnapShot_Click
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   Š   NetEye Files| *.nef|NetEye Key Files | *.nkf|NetEye Image Files|*.nif|NetEye BMP Files | *.bmp| All NetEye Files| *.nef;*.nkf;*.nif; *.bmp                               