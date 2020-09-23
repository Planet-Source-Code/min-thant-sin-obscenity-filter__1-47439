VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ObscenityFilter"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1410
   ControlBox      =   0   'False
   Icon            =   "ObscenityFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   1410
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   75
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   450
      Picture         =   "ObscenityFilter.frx":08CA
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////
'/// Obscenity Filter by Min Thant Sin                                               /////
'/// Version 1.0, Wednesday, August 6, 2003                                     /////
'/// Comments and suggestions are welcome.                                       /////
'/// Any bugs? Feel free to e-mail me at < minsin999@hotmail.com > /////
'////////////////////////////////////////////////////////////////////

'*********** My constants ********************
Private Const IE_CLASS_NAME = "IEFRAME"

'*********** My variables *********************
Private DirtyWords() As String

'*********** Win32 API constants & declarations *********************
Private Const GW_CHILD = 5
Private Const GW_NEXT = 2
Private Const WM_CLOSE = &H10

'Private Const WM_SYSCOMMAND = &H112  'For closing IE window using SendMessage
'Private Const SC_CLOSE = &HF060& 'For closing IE window using SendMessage
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'This sub loads the words from WordsList.obf in the App path
Private Sub LoadWordsList()
      On Error GoTo ErrHandler
      
      Dim Words As String
      Dim intFileFree As Integer
      Dim I As Integer
      
      Open App.Path & "\WordsList.obf" For Input As #1
      
      I = 0
      
      'Load the dirty words into array
      Do Until EOF(1)
            'Input line by line and store it in variable Words
            Line Input #1, Words
            I = I + 1
            
            ReDim Preserve DirtyWords(1 To I)
            'Store it in the array
            DirtyWords(I) = Words
            
            DoEvents
      Loop
      
      Close #1
      
      Exit Sub
ErrHandler:
      MsgBox Err.Description
      On Error Resume Next
      
      'Write error message to file if there is any
      intFileFree = FreeFile()
      Open App.Path & "\ErrorLog.log" For Output As #intFileFree
            Print #1, Err.Description
      Close #intFileFree
End Sub

Private Sub CloseDirtyWindows()
      Dim hndWindow As Long
      Dim retVal As Long
      Dim nMaxCount As Integer
      Dim I As Integer
      Dim lpClassName As String
      Dim lpCaption As String
      
      nMaxCount = 256
      
      'Get the first child of desktop window
      hndWindow = GetWindow(GetDesktopWindow(), GW_CHILD)
                  
      'Find the rest siblings of that child window
      Do While hndWindow <> 0
            'We don't check windows that are not visible
            retVal = IsWindowVisible(hndWindow)
            
            If retVal <> 0 Then    '// Main - If
                  'Create buffers to retrieve class name & window text
                  lpClassName = String(nMaxCount, Chr(0))
                  lpCaption = String(nMaxCount, Chr(0))
                  
                  'Get the class name of the window
                  retVal = GetClassName(hndWindow, lpClassName, nMaxCount)
                  lpClassName = Left(lpClassName, retVal)
                  
                  'Check if it is IEFrame
                  If UCase(Left(lpClassName, 7)) = IE_CLASS_NAME Then
                        'Get the caption of the window
                        retVal = GetWindowText(hndWindow, lpCaption, nMaxCount)
                        lpCaption = Left(lpCaption, retVal)
                        
                        'Check for obscene words in window's caption
                        For I = 1 To UBound(DirtyWords())
                              If InStr(1, lpCaption, DirtyWords(I), vbTextCompare) > 0 Then
                                    'Close that window
                                    PostMessage hndWindow, WM_CLOSE, 1, 0
                                    'SendMessage hndWindow, WM_SYSCOMMAND, SC_CLOSE, 0
                                    Exit For
                              End If
                        Next I
                  End If
                  
            End If                   '// Main - End If
            
            'Get next window
            hndWindow = GetWindow(hndWindow, GW_NEXT)
            DoEvents
      Loop
            
End Sub

Private Sub Form_Load()
      'Don't allow two instances
      If App.PrevInstance Then End
      
      App.TaskVisible = False
      LoadWordsList
      Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      'The operating system is shutting down, so end the app
      If UnloadMode = vbAppWindows Then
            'End frees all memory allocated for this app
            End
      End If
End Sub

'Interval is set to 3000 (three seconds)
Private Sub Timer1_Timer()
      'NOTE:
      '(1) To END the application, create (if there is none)
      '      or rename the text file to "StopFilter.txt"
      '(2) To run the application without terminating itself
      '      after Timer1's interval has elapsed, rename the
      '      text file to "StopFilter().txt" or something like that.
      
      If Dir(App.Path & "\StopFilter.txt") <> "" Then End
      CloseDirtyWindows
End Sub
