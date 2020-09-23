VERSION 5.00
Begin VB.Form frmSlave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSub.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmSub.frx":000C
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgScroll 
      Height          =   285
      Left            =   270
      Top             =   615
      Width           =   3945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here!"
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
      Left            =   1815
      TabIndex        =   1
      Top             =   630
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Put extra controls in me!"
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
      Left            =   1230
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
End
Attribute VB_Name = "frmSlave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Used to store the location of the original WndProc
Public OrigWndProc As Long
'Used to store the current vertical offset (in pixels) of the "slave" window
'relative to the bottom of the "master" window.
Public YOffset As Long
'Used to store the handle to the region that the form will use for display
Private hRgn As Long

Private Sub Form_Activate()

'    'This will prevent the window from jumping in front of our main window.
'    'Without this line, we don't get the correct visual effect.
'    Call SetWindowPos(Me.hWnd, frmMaster.hWnd, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'
End Sub

Private Sub Form_Load()

    Dim hRgnTemp As Long
    
    'Initialize the YOffset value to -59 so it's neatly tucked away.
    'This value is usually the negative of 1 pixel less than the total
    'height of the region defining the shape of the form (in this case 60 pixels).
    YOffset = -59
    
    'We need to shape the window to our needs (rectangles are so boring).
    'We'll start by creating a rectangle region smaller than the whole window.
    hRgn = CreateRectRgn(20, 0, 279, 50)
    'Now we'll add a round rectangle region at the bottom (this will be the
    'window's handlebar).
    hRgnTemp = CreateRoundRectRgn(18, 40, 282, 60, 15, 15)
    'Now we combine the regions into a new single region.
    'RGN_OR simply adds the two regions together.
    Call CombineRgn(hRgn, hRgn, hRgnTemp, RGN_OR)
    'We don't need this region any longer so we can delete it now
    'We can't delete the other region yet because the window still needs it
    'during redraws.
    Call DeleteObject(hRgnTemp)
    
    'Assign the newly created region to the window
    Call SetWindowRgn(Me.hWnd, hRgn, True)

    'Subclass the window so we can monitor the WM_MOVE message.
    'See the WndCallbackProc function for more details.
    OrigWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndCallbackProc)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'UnSubClass the window so we don't get a crash.
    Call SetWindowLong(Me.hWnd, GWL_WNDPROC, OrigWndProc)

    'Clear the window region so we don't have redraw problems.
    'This problem is particularly noticeable on a Windows 95 machine.
    Call SetWindowRgn(Me.hWnd, 0&, True)
    'Delete the region object since we no longer need it.
    'Not deleting it doesn't cause problems but leaves resources open that are
    'never reclaimed (very messy).
    Call DeleteObject(hRgn)

End Sub

Private Sub imgScroll_Click()

    'This little code snippit will scroll the form smoothly down or up as necessary.

    If (YOffset = -20) Then 'the form is already down, scroll it up
        While YOffset > -59
            'We'll move the form 1 pixel at a time to make it smooth.
            frmSlave.Move frmSlave.Left, frmSlave.Top - Screen.TwipsPerPixelY
            'Make sure we keep the YOffset value synchronized
            YOffset = YOffset - 1
            'Allow the system to redraw the screen. It looks ugly without this.
            DoEvents
        Wend
    Else 'the form is tucked away, scroll it down
        While YOffset < -20
            frmSlave.Move frmSlave.Left, frmSlave.Top + Screen.TwipsPerPixelY
            YOffset = YOffset + 1
            DoEvents
        Wend
    End If

End Sub
