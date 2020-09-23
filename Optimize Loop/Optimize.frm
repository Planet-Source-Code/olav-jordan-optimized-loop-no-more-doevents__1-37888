VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                                 'looks at message and removes/leaves it if there is one
                                 'returns nonzero if a message was in event queue
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
                                 'dispatches message calls the right message handling procedure
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
                                 'virtual accelerator key translator
                                 'dont worry about what it does just leave it there
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
                                 'holds elapsed time since windows was started
Private Declare Function GetTickCount Lib "kernel32" () As Long
                                 'gets next message in event queue
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
                                 'checks if there is a message in event queue
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

                           
Private Const MY_WM_QUIT = &HA1     'WM_QUIT in api viewer is wrong this is the right constant
                                       
Private Const PM_REMOVE = &H1       'paramater on peekmessage to remove or leave message in queue
Private Const PM_NOREMOVE = &H0
                                    'type of events that can happen with window
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSEMOVE = &H2
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8  'any other message
Private Const QS_TIMER = &H10
Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

'extra messages that can be sent (not used in example)
Private Const QS_SENDMESSAGE = &H40    'message sent by other thread or app
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
'*************************

Private Type POINTAPI
   X As Long
   y As Long
End Type

Private Type MSG
   hwnd     As Long        'window where message occured
   Message  As Long        'message id itself
   wParam   As Long        'further defines message
   lParam   As Long        'further defines message
   time     As Long        'time of message event
   pt       As POINTAPI    'position of mouse
End Type

Dim Message As MSG         'holds message recieved from queue

Private Sub Form_Load()
   Dim HoldTime As Long    'holds time
   Dim I As Long           'counter
      
   Me.Show
   HoldTime = GetTickCount
   
   Do While GetTickCount - HoldTime < 1000         'runs loop for 100 milliseconds or 1 second
''**************************doevent1************************** fastest
'      If PeekMessage(Message, 0, 0, 0, PM_REMOVE) Then        'checks for a message in the queue and removes it if there is one
'         TranslateMessage Message                             'translates the message(dont need if there is no menu)
'
'         DispatchMessage Message                              'dispatches the message to be handled
'
'         If (Message.Message = MY_WM_QUIT) Then               'if the message is to quit then exit the loop
'            Exit Do
'         End If
'      End If
''*************************************************************

''**************************doevent2************************** third fastest
'      If PeekMessage(Message, 0, 0, 0, PM_NOREMOVE) Then      'checks for a message in the queue
'         DoEvents                                             'dispatches any messages in the queue
'      End If
''************************************************************

''**************************doevent3************************** second fastest
'      If PeekMessage(Message, 0, 0, 0, PM_NOREMOVE) Then      'checks for a message in the queue
'         If GetMessage(Message, 0, 0, 0) = 0 Then Exit Do     'check for error
'         DispatchMessage Message                              'dispatch message
'
'         If Message.Message = MY_WM_QUIT Then                 'if the message is to quit then exit the loop
'            Exit Do
'         End If
'      Else
'                    'main part of program goes here instead of outside if statement
'      End If
''************************************************************

''**************************doevent4************************** fourth fastest
'      If GetQueueStatus(QS_ALLEVENTS) Then                    'checks the queue for messages
'         DoEvents                                             'dispatches all the messages
'      End If
''************************************************************

''**************************doevent5************************** slowest
'      DoEvents                                                'dispatches messages in the queue
''************************************************************

      I = I + 1
   Loop
      
   'becareful what you put here after loop
   
   Print I  'this will return error if program is closed before the 1 second loop timer
            'because the window is already unloaded
End Sub


