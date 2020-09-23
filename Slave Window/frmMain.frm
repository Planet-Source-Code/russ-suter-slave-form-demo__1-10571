VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Slave Form Demo..."
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1965
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7854
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":792C
            Key             =   "CloseActive"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A04
            Key             =   "Minimize"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AD0
            Key             =   "MinimizeActive"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B9C
            Key             =   "Maximize"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbTitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4425
      Begin VB.PictureBox pbMaximize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   180
         Left            =   3900
         Picture         =   "frmMain.frx":7C68
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   180
      End
      Begin VB.PictureBox pbMinimize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   180
         Left            =   3720
         Picture         =   "frmMain.frx":7D24
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   180
      End
      Begin VB.PictureBox pbClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   180
         Left            =   4110
         Picture         =   "frmMain.frx":7DE0
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Put controls in me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1470
      TabIndex        =   4
      Top             =   480
      Width           =   1635
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Used to store the location of the original WndProc
Public OrigWndProc As Long
'Used to store the handle to the region that the form will use for display
Private hRgn As Long
Private hRgnPic As Long

Private Sub Form_Load()

    Dim hRgnTemp As Long
    
    'Show the form in the TaskBar even though it's borderless
    Call SetWindowLong(Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_SYSMENU)

    'We need to shape the window to our needs (rectangles are so boring).
    'We'll start with a basic round rectangle region (round out the corners).
    hRgn = CreateRoundRectRgn(0, 0, 300, 100, 25, 25)
    'We'll add a cutout for the "slave" form to show through.
    hRgnTemp = CreateRoundRectRgn(20, 80, 280, 125, 15, 15)
    'Now we combine the regions into a new single region.
    'RGN_XOR will mask out any areas where the regions overlap.
    Call CombineRgn(hRgn, hRgn, hRgnTemp, RGN_XOR)
    'We don't need this region any longer so we can delete it now.
    'We can't delete the other region yet because the window still needs it
    'during redraws.
    Call DeleteObject(hRgnTemp)
    
    'Assign the newly created window region to the window.
    Call SetWindowRgn(Me.hWnd, hRgn, True)
    
    'Round the corners of our title bar to match the window
    hRgnPic = CreateRoundRectRgn(0, 0, 295, 30, 24, 24)
    Call SetWindowRgn(pbTitleBar.hWnd, hRgnPic, True)
    
    'Subclass the window so we can monitor the WM_MOVE message.
    'See the WndCallbackProc function for more details.
    OrigWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndCallbackProc)

    Call DrawTitleBar(True)

    'Modify the system menu to disable the "Move" item since we can't respond
    'to this command correctly with a borderless window.
    Call ModifyMenu(GetSystemMenu(Me.hWnd, False), SC_MOVE, MF_BYCOMMAND Or MF_GRAYED, 0&, "&Move")

    'Show the "slave" window (it will position itself automatically).
    frmSlave.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'UnSubClass the window so we don't get a crash.
    Call SetWindowLong(Me.hWnd, GWL_WNDPROC, OrigWndProc)

    'Clear the window regions so we don't have redraw problems.
    'This problem is particularly noticeable on a Windows 95 machine.
    Call SetWindowRgn(Me.hWnd, 0&, True)
    Call SetWindowRgn(pbTitleBar.hWnd, 0&, True)
    'Delete the region objects since we no longer need them.
    'Not deleting them doesn't cause problems but leaves resources open that are
    'never reclaimed (very messy).
    Call DeleteObject(hRgn)
    Call DeleteObject(hRgnPic)

End Sub

Private Sub pbClose_Click()

    'Exiting the program. We cannot use "End" here because of the subclassing
    'so we just unload all the forms.
    Unload frmSlave
    Unload frmMaster

End Sub

Private Sub pbClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Change the picture so it looks like the button is being pressed
    Set pbClose.Picture = ImageList.ListImages("CloseActive").Picture

End Sub

Private Sub pbClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Here we handle the case where the mouse moves outside the boundaries
    'of the picture box control.
    If (Button <> 0) Then
        If ((X < 0) Or (Y < 0) Or (X > pbClose.ScaleWidth) Or (Y > pbClose.ScaleHeight)) Then
            Set pbClose.Picture = ImageList.ListImages("Close").Picture
        Else
            Set pbClose.Picture = ImageList.ListImages("CloseActive").Picture
        End If
    End If

End Sub

Private Sub pbClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Change the picture so it looks like the button has been released
    Set pbClose.Picture = ImageList.ListImages("Close").Picture
    'Allow the system to redraw the button before performing operations
    'specified in the Click event.
    DoEvents

End Sub

Private Sub pbMinimize_Click()

    'This method is the same method used by the standard minimize button.
    'The Form.WindowState = vbMinimized method is VB's own method.
    'If the user has a sound associated with the minimize action, this method
    'will play it. VB's method will not.
    Call SendMessage(frmMaster.hWnd, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&)

End Sub

Private Sub pbMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Change the picture so it looks like the button is being pressed
    Set pbMinimize.Picture = ImageList.ListImages("MinimizeActive").Picture

End Sub

Private Sub pbMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Here we handle the case where the mouse moves outside the boundaries
    'of the picture box control.
    If (Button <> 0) Then
        If ((X < 0) Or (Y < 0) Or (X > pbMinimize.ScaleWidth) Or (Y > pbMinimize.ScaleHeight)) Then
            Set pbMinimize.Picture = ImageList.ListImages("Minimize").Picture
        Else
            Set pbMinimize.Picture = ImageList.ListImages("MinimizeActive").Picture
        End If
    End If

End Sub

Private Sub pbMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Change the picture so it looks like the button has been released
    Set pbMinimize.Picture = ImageList.ListImages("Minimize").Picture
    'Allow the system to redraw the button before performing operations
    'specified in the Click event.
    DoEvents

End Sub

Private Sub pbTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Since we have modified the system menu, this little trick won't work.
    'Therefore, we need to reset the menu before we call it. WndCallbackProc will
    'watch for the WM_EXITSIZEMOVE message so we can disable the menu again when
    'the user has finished moving the form. The result is completely transparent
    'to the user.
    Call GetSystemMenu(Me.hWnd, True)

    'This is so we can move the window without a title bar (it's an old trick).
    'We need to release the mouse capture back to the system so it can track movement.
    ReleaseCapture
    'Now we just send a message simulating a MouseDown on a title bar.
    Call SendMessage(frmMaster.hWnd, WM_NCLBUTTONDOWN, 2, 0&)

End Sub

Public Sub DrawTitleBar(ByVal DrawActive As Boolean)

    'This function will fill our custom titlebar with the same colors as
    'standard windows. It will even do gradients for 98 or 2000.
    
    Dim X As Long
    Dim FillColor As Long
    Dim rDiff As Integer, gDiff As Integer, bDiff As Integer
    Dim rLeft As Integer, gLeft As Integer, bLeft As Integer
    Dim rFill As Integer, gFill As Integer, bFill As Integer
    Dim Percentage As Single
    
    'We need to find out what 2 colors are used for the gradient fill
    
    'First we'll find the right color (color 2)
    If DrawActive Then
        FillColor = GetSysColor(COLOR_GRADIENTACTIVECAPTION)
    Else
        FillColor = GetSysColor(COLOR_GRADIENTINACTIVECAPTION)
    End If
    'Now separate the value (returned as a COLORREF) into R,G,B components
    rFill = LSByte(FillColor)
    gFill = NLSByte(FillColor)
    bFill = NMSByte(FillColor)
    
    'Now find the left color (color 1)
    If DrawActive Then
        FillColor = GetSysColor(COLOR_ACTIVECAPTION)
    Else
        FillColor = GetSysColor(COLOR_INACTIVECAPTION)
    End If
    'Now separate the value (returned as a COLORREF) into R,G,B components
    rLeft = LSByte(FillColor)
    gLeft = NLSByte(FillColor)
    bLeft = NMSByte(FillColor)
    
    If ((OSType() <> 98) And (OSType() <> 5)) Then
        'If we're not using Windows 98 or 2000, gradient fill isn't supported.
        'We'll just set the backcolor of the picture box.
        pbTitleBar.BackColor = FillColor
    Else
        'Find out the difference between the 2 color values and store them
        'in separate variables.
        rDiff = rFill - rLeft
        gDiff = gFill - gLeft
        bDiff = bFill - bLeft
        
        'Make sure we're using pixels or the gradient fill won't look right.
        pbTitleBar.ScaleMode = vbPixels
        
        'loop once for each pixel (width) of the picture box
        For X = 0 To pbTitleBar.ScaleWidth
            Percentage = Round((X / pbTitleBar.ScaleWidth), 2)
            pbTitleBar.Line (X, 0)-(X, pbTitleBar.ScaleHeight), FillColor
            rFill = rLeft + rDiff * Percentage
            gFill = gLeft + gDiff * Percentage
            bFill = bLeft + bDiff * Percentage
            FillColor = RGB(rFill, gFill, bFill)
        Next
    End If
    
    'Now draw the title of the window
    pbTitleBar.CurrentX = 10
    pbTitleBar.CurrentY = 0
    If DrawActive Then
        pbTitleBar.ForeColor = GetSysColor(COLOR_CAPTIONTEXT)
    Else
        pbTitleBar.ForeColor = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    End If
    pbTitleBar.Print Me.Caption
    
End Sub
