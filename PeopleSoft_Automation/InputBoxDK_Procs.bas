Attribute VB_Name = "InputBoxDK_Procs"
Option Explicit
 
 '////////////////////////////////////////////////////////////////////
 'Password masked inputbox
 'Allows you to hide characters entered in a VBA Inputbox.
 '
 'Code written by Daniel Klann
 'http://www.danielklann.com/
 'March 2003
 
 '// Kindly permitted to be amended
 '// Amended by Ivan F Moala
 '// http://www.xcelfiles.com
 '// April 2003
 '// Works for Xl2000+ due the AddressOf Operator
 '////////////////////////////////////////////////////////////////////
 
 '********************   CALL FROM FORM *********************************
 '    Dim pwd As String
 '
 '    pwd = InputBoxDK("Please Enter Password Below!", "Database Administration Security Form.")
 '
 '    'If no password was entered.
 '    If pwd = "" Then
 '        MsgBox "You didn't enter a password!  You must enter password to 'enter the Administration Screen!" _
 '        , vbInformation, "Security Warning"
 '    End If
 '**************************************
 
 
 
 'API functions to be used
#If VBA7 Then
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As LongPtr) As Long
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr
#Else
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#End If

'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0
 
Private hHook As Long
 
Public Function NewProc(ByVal lngCode As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
     
    Dim RetVal
    Dim strClassName As String, lngBuffer As Long
     
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If
     
    strClassName = String$(256, " ")
    lngBuffer = 255
     
    If lngCode = HCBT_ACTIVATE Then 'A window has been activated
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        If Left$(strClassName, RetVal) = "#32770" Then 'Class name of the Inputbox
             'This changes the edit control so that it display the password character *.
             'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
    End If
     
     'This line will ensure that any other hooks that may be in place are
     'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
     
End Function
 
 '// Make it public = avail to ALL Modules
 '// Lets simulate the VBA Input Function
Public Function InputBoxDK(Prompt As String, Optional Title As String, _
    Optional default As String, _
    Optional Xpos As Long, _
    Optional Ypos As Long, _
    Optional Helpfile As String, _
    Optional Context As Long) As String
     
    Dim lngModHwnd As LongPtr, lngThreadID As LongPtr
    
     '// Lets handle any Errors JIC! due to HookProc> App hang!
    On Error GoTo ExitProperly
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
     
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    If Xpos Then
        InputBoxDK = InputBox(Prompt, Title, default, Xpos, Ypos, Helpfile, Context)
    Else
        InputBoxDK = InputBox(Prompt, Title, default, , , Helpfile, Context)
    End If
     
ExitProperly:
    UnhookWindowsHookEx hHook
     
End Function
 
Private Sub TestDKInputBox()
    Dim x
     
    x = InputBoxDK("Type your password here.", "Password Required")
    If x = "" Then End
    If x <> "yourpassword" Then
        MsgBox "You didn't enter a correct password."
        End
    End If
     
    MsgBox "Welcome Creator!", vbExclamation
     
End Sub

