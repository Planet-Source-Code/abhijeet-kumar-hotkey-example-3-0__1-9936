Attribute VB_Name = "modHotkey"
'Program:    Hotkey Example 3.00
'Created by: Abhijeet Kumar
'Created on: Tuesday, July 25th, 2000

'This project was done in order to help you better understand
'the many uses and implementations of subclassing. Please keep
'in mind that this example, although fully functional and
'debugged, is hardly a valid reason for subclassing.

'You are free to use and modify this source, but please, if
'you do so, give me credit somehow.

Option Explicit  'Force the declaration of ALL variables

Public Hot_Key&  'The Public variable used to store the first
                 'key of the Hotkey combination

Public Hot_Atom% 'The Public variable used to store the Hotkey
                 'Atom number

Public Hot_hWnd& 'The Public variable used to store the hWnd
                 'used during the subclass process

Public Hot_Letter As KeyCodeConstants
                 'The Public variable used to store the second
                 'part of the Hotkey combination

'Begin declaration of Public API Constants

Public Const GWL_WNDPROC = (-4)
Public Const WM_HOTKEY = &H312
Public Const WM_NCDESTROY = &H82
Public Const MOD_ALT = &H1
Public Const MOD_SHIFT = &H4
Public Const MOD_CONTROL = &H2

'Begin declaration of API Functions

Declare Function CallWindowProc& Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Declare Function RegisterHotKey& Lib "User32" (ByVal hWnd&, ByVal id&, ByVal fsModifiers&, ByVal vk&)
Declare Function UnRegisterHotKey& Lib "User32" (ByVal hWnd&, ByVal id&)
Declare Function SetWindowLong& Lib "User32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
Declare Function GetWindowLong& Lib "User32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&)
Declare Function GlobalAddAtom% Lib "Kernel32" Alias "GlobalAddAtomA" (ByVal lpString$)
Declare Function GlobalDeleteAtom% Lib "Kernel32" (ByVal nAtom%)

Sub Hotkey_Kill()
Dim Buffer&
On Error Resume Next

'This Sub will stop the subclassing process and restore the
'normal, program defaults.

    If Hot_hWnd <> 0 Then 'If an hWnd has been assigned to
                          'receive the subclassed messages,
                          'then...

        Buffer = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, _
            Hot_hWnd)
            'Unhook the Main Form hWnd from receiving the
            'subclassed messages.

        Buffer = UnRegisterHotKey(frmMain.hWnd, Hot_Atom)
            'Unregister the Hotkey.

        Buffer = GlobalDeleteAtom(Hot_Atom)
            'Unregister the assigned Hotkey Atom.

        Hot_hWnd = 0
            'Devalue hWnd variable so as to remove all
            'evidence that a window was hooked to receive
            'subclassed messages in the first place.

    End If 'End If routine

End Sub

Sub Hotkey_Update()
Dim Buffer&
On Error Resume Next

'This Sub will register a new Hotkey combination for
'subclassing.

    Hotkey_Kill
        'Remove previous Hotkey combination and subclassing.

    Hot_Letter = Asc(frmMain.cboLetter.Text)
        'Convert the Combo box selection to an ASCII value.

    Hot_Atom = GlobalAddAtom("NewMacro")
        'Assign an Atom to the Hotkey combination.

    Buffer = RegisterHotKey(frmMain.hWnd, Hot_Atom, _
        Hot_Key, Hot_Letter)
        'Register the actual Hotkey combination using the
        'values stored in the Public variables.

    Hot_hWnd = GetWindowLong(frmMain.hWnd, GWL_WNDPROC)
        'Force the Program to receive the subclassed messages.

    Buffer = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, AddressOf Hotkey_Used)
        'Too complicated to explain. Please cite an advanced
        'Visual Basic reference book.

End Sub

Function Hotkey_Used&(ByVal hWnd&, ByVal Msg&, ByVal WP&, ByVal LP&, Result&)
On Error Resume Next

'When a possible Hotkey combination is pressed, this Function
'is used to validify the message received.

    Select Case Msg 'Begin Select routine

        Case WM_NCDESTROY 'If the message received is
                          'WM_NCDESTROY, stop subclassing the
                          'specified Hotkey.

            Hotkey_Kill 'Stop subclassing Hotkey

        Case WM_HOTKEY 'If the message received is WM_HOTKEY,
                       'then the Hotkey combination has been
                       'pressed.

            If WP = Hot_Atom Then Hotkey_Execute
                'If the WP parameter is equal to the Atom
                'number of the Hotkey, then execute
                'user-specified command.

        Case Else 'Begin counter Select routine

            Hotkey_Used = CallWindowProc(Hot_hWnd, frmMain.hWnd, Msg, WP, LP)
                'Assign the function an API call value.

    End Select 'End Select routine

End Function

Sub Hotkey_Execute()
Dim Buffer#
On Error Resume Next

'When the Hotkey combination is pressed, the coding within
'this Sub is executed. It can be ANYTHING valid.

    Buffer = Shell("c:\windows\notepad.exe", vbNormalFocus)
        'Run the Windows Notepad executable file with focus.

End Sub
