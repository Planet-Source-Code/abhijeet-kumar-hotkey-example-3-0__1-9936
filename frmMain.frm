VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hotkey Example 3.0"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLetter 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Subclass desired LETTER"
      Top             =   360
      Width           =   735
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Shift"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Subclass the SHIFT key"
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Control"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Subclass the CONTROL key"
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Alt"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Subclass the ALT key"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "&Enable Hotkey"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Force the declaration of ALL variables

Sub Form_Enabled(Buffer As Boolean)

'This is just a simple way of disabling all of the controls
'on the Form quickly, efficiently, and without any redundant
'coding.

    optKey(0).Enabled = Buffer  'The ALT radio button
    optKey(1).Enabled = Buffer  'The CONTROL radio button
    optKey(2).Enabled = Buffer  'The SHIFT radio button

    cmdMain(1).Enabled = Buffer 'The Enable/Disable button
    cboLetter.Enabled = Buffer  'The Exit button

End Sub

Private Sub Form_Load()
Dim Index%
On Error Resume Next

'In this procedure, we want to load both our public variables
'and combo box with their respective default information. The
'combo box receives a static list of the capital letter
'alphabet, A through Z, and the public variables receive their
'initial key values.

    Hot_Key = MOD_ALT   'Load the ALT key
    Hot_Letter = vbKeyA 'Load the A key

    For Index = 65 To 90: DoEvents   'Begin For-Loop

        cboLetter.AddItem Chr(Index) 'Add the captial letter
                                     'to Combo List
    Next Index                       'Loop until finished

End Sub

Private Sub optKey_Click(Index As Integer)

'Since an array was used for the Radio buttons, assigning
'the correct values and reducing redundant coding becomes
'almost automatic.

    Select Case Index 'Begin Select routine

        Case 0: Hot_Key = MOD_ALT
            'If the ALT Radio button was selected, assign the
            'ALT key to the Hot_Key public variable.

        Case 1: Hot_Key = MOD_CONTROL
            'If the CONTROL Radio button was selected, assign
            'the CONTROL key to the Hot_Key public variable.

        Case 2: Hot_Key = MOD_SHIFT
            'If the SHIFT Radio button was selected, assign
            'the SHIFT key to the Hot_Key public variable.

    End Select 'End Select routine

End Sub

Private Sub cmdMain_Click(Index As Integer)

'Since an array was used for the Command buttons, assigning
'the correct values and reducing redundant coding becomes
'almost automatic.

    Select Case Index 'Begin Select routine

        Case 0 'If the Enable/Disable button was clicked
               'then...

            If cmdMain(0).Caption = "&Enable Hotkey" Then
                'If the Command button caption remains exactly
                'the same as when the program was run, then
                'that means that it most likely hasn't been
                'clicked as of yet, and that no Hotkey is
                'being subclasses at the moment. So...

                Form_Enabled False: Hotkey_Update
                    'Disable the Form controls so as to
                    'prevent any tampering with the gathered
                    'data, and begin subclassing the selected
                    'Hotkey combination.

                cmdMain(0).Caption = "&Disable Hotkey"
                    'Change the Command button caption so as
                    'to notify the program later on that the
                    'key has been pressed, and that a Hotkey
                    'combination is being subclasses at the
                    'moment.

            Else 'Begin the counter If routine

                Form_Enabled True: Hotkey_Kill
                    'The Command button has been clicked
                    'again, notifying the program that the
                    'user wishes to cease the Hotkey
                    'subclassing and restore the Form default.

                cmdMain(0).Caption = "&Enable Hotkey"
                    'Restore the Command button caption back
                    'to its default value, thereby removing
                    'any evidence a Hotkey was ever subclassed
                    'in the first place.

            End If 'End If routine

        Case 1 'If the Exit button was clicked then...

            Unload Me 'Go to the Unload_Form procedure

    End Select 'End Select routine

End Sub

Private Sub Form_Unload(Cancel As Integer)

'The purpose of this procedure is to avoid ending the program
'while a Hotkey combination is still being subclassed. Why?
'Because in more cases then not, an Illegal Operation error
'will be produced.

    Hotkey_Kill 'Stop subclassing the Hotkey

End Sub
