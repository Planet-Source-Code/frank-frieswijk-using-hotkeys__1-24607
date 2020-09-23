VERSION 5.00
Begin VB.Form frmHOTKEY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create a Hotkey"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMAIN 
      Height          =   2715
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   2940
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   945
         TabIndex        =   11
         Top             =   945
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   3
         Left            =   945
         TabIndex        =   10
         Top             =   1170
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   4
         Left            =   945
         TabIndex        =   9
         Top             =   1395
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   5
         Left            =   945
         TabIndex        =   8
         Top             =   1620
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   6
         Left            =   945
         TabIndex        =   7
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   7
         Left            =   1935
         TabIndex        =   6
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   8
         Left            =   1935
         TabIndex        =   5
         Top             =   945
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   9
         Left            =   1935
         TabIndex        =   4
         Top             =   1170
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   10
         Left            =   1935
         TabIndex        =   3
         Top             =   1395
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   11
         Left            =   1935
         TabIndex        =   2
         Top             =   1620
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   12
         Left            =   1935
         TabIndex        =   1
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox chkKEY 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   12
         Top             =   720
         Width           =   195
      End
      Begin VB.CommandButton cmdQUIT 
         Caption         =   "Exit program"
         Height          =   375
         Left            =   855
         TabIndex        =   26
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F1"
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
         Index           =   0
         Left            =   630
         TabIndex        =   25
         Top             =   720
         Width           =   225
      End
      Begin VB.Label lblENABLE 
         AutoSize        =   -1  'True
         Caption         =   "Enable function key:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   405
         TabIndex        =   24
         Top             =   270
         Width           =   2145
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F2"
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
         Index           =   2
         Left            =   630
         TabIndex        =   23
         Top             =   945
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F3"
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
         Index           =   3
         Left            =   630
         TabIndex        =   22
         Top             =   1170
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F4"
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
         Index           =   4
         Left            =   630
         TabIndex        =   21
         Top             =   1395
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F5"
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
         Index           =   5
         Left            =   630
         TabIndex        =   20
         Top             =   1620
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F6"
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
         Index           =   6
         Left            =   630
         TabIndex        =   19
         Top             =   1845
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F12"
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
         Index           =   7
         Left            =   1530
         TabIndex        =   18
         Top             =   1845
         Width           =   330
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F11"
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
         Index           =   8
         Left            =   1530
         TabIndex        =   17
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F10"
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
         Index           =   9
         Left            =   1530
         TabIndex        =   16
         Top             =   1395
         Width           =   330
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F9"
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
         Index           =   10
         Left            =   1530
         TabIndex        =   15
         Top             =   1170
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F8"
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
         Index           =   11
         Left            =   1530
         TabIndex        =   14
         Top             =   945
         Width           =   225
      End
      Begin VB.Label lblKEY 
         AutoSize        =   -1  'True
         Caption         =   "F7"
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
         Index           =   12
         Left            =   1530
         TabIndex        =   13
         Top             =   720
         Width           =   225
      End
   End
   Begin VB.Label Label3 
      Caption         =   "ICQ: 75949744"
      Height          =   375
      Left            =   1200
      TabIndex        =   29
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "SpeedX@hotmail.com"
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "[ Contacts ]"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmHOTKEY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'We will use the HotKey variable to pass the hotkey/functionkey
'value to enable
Dim HotKey As Byte

'For the unload part we will need to know if the hotkey
'was enabled or not (there are 12 possible hotkeys)
Dim HotKeyEnabled(12) As Boolean

'If the user clicked for the first time let them know what
'to do by displaying a messagebox only one time
Dim ClickedBefore As Boolean

Public Sub Cleanup()

For HotKey = 1 To 12
    If HotKeyEnabled(HotKey) = True Then
        Call DeleteHotkey
    End If
Next HotKey

Unload Me

'Don't end with End since this will cause the program to crash
'even if you restored the SetWindowLong property to the OldwndHnd

End Sub


Public Sub CreateHotkey()

Dim ReturnValue As Long

'The Hotkey will need an ID since there are several
'hotkeys you will want know which one has been pressed

'To keep it simple, we will keep the ID number the same
'as the Function key number that has been selected

HotKeyID(HotKey) = HotKey

'There are 12 function keys. The Ascii number for
'Function-Key 1 = 112
'So if we add 111 to our Hotkey number we know the value
'of the function key that will be the hotkey

HotKey = HotKey + 111

'Now that we know the value and ID of the key to register as a
'HotKey we can actually register the key as a hotkey

'(note: since we added 111 to HotKey we will need to subtract it again
' to get the right HotKey number for the ID)
ReturnValue = RegisterHotKey(hWnd, HotKeyID(HotKey - 111), 0, HotKey)



End Sub

Public Sub DeleteHotkey()

Dim ReturnValue As Long

'To disable/unload the selected hotkey (index number of the checkbox
'that was clicked on) simply unregister the HotKeyID (from the form it
'was registered to)
ReturnValue = UnregisterHotKey(hWnd, HotKeyID(HotKey))

End Sub


Public Sub SetToolTips()

Dim i As Byte
For i = 1 To 12
    chkKEY(i).ToolTipText = "If a key has been enabled, press it to see what happens"
Next i

End Sub

Private Sub chkKEY_Click(Index As Integer)

'A checkbox was clicked. We need to find out if now it is
'checked or not and act accordingly

If chkKEY(Index).Value = 1 Then 'The box that was clicked is now checked
    'so enable the function key as a hotkey.
    'We know that index holds the value of the function-key to enable
    HotKey = Index
    Call CreateHotkey 'Go to the CreateHotKey sub to actually create the hotkey
    HotKeyEnabled(Index) = True
Else 'The box is now unchecked, so we need to disable this hotkey
    HotKey = Index
    Call DeleteHotkey
    HotKeyEnabled(Index) = False
End If

If ClickedBefore = False Then
    MsgBox "Now you can press F" & Index & " to see that it acts as a hotkey." & vbCrLf & _
        "If you want you can select more hotkeys.", vbInformation, App.Title
    ClickedBefore = True
End If

End Sub

Private Sub cmdQUIT_Click()

Call Cleanup

End Sub

Private Sub Form_Load()

'Because I forgot to add the tooltip I will do this first (...)
Call SetToolTips

'In order to be able to see if the hotkey has been pressed
'we need to intercept all window messages and find out if
'a message contains the hotkey information

'In order to intercept all window messages we redirect all
'messages to a function called WindowProc.
'The WindowProc can be found in the modHotkey module.

'Using functions called by Windows are called Callbacks.
'Callbacks (their declarations and constants) always need
'to be in a module
'Intercepting messages like this is called Subclassing

'The line below tells the current form to redirect all window
'messages to a custom procedure handler, in this case (and usually)
'called WindowProc.

OldwndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Private Sub Form_Terminate()

Call Cleanup

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call Cleanup

End Sub


