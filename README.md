<div align="center">

## Optimized loop \(no more doevents\)


</div>

### Description

Although this is the 3rd post of the same type of

thing i would like to submit this anyway. John G.

posted a better way of using doevents that does

speed up the loop by about 100% and JohnB entered

a class to do this based on John G.'s article. I have desided to give alternatives to using doevents all together or a slightely faster way

of using doevents with an example showing the

speed of each that is included in the zip. I ask that you DON'T vote for my code but comments are welcome.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-08-12 19:22:42
**By**             |[Olav Jordan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/olav-jordan.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Optimized\_1173138122002\.zip](https://github.com/Planet-Source-Code/olav-jordan-optimized-loop-no-more-doevents__1-37888/archive/master.zip)





### Source Code

```
If PeekMessage(Message, 0, 0, 0, PM_REMOVE) Then
    TranslateMessage Message
    DispatchMessage Message
    If (Message.Message = MY_WM_QUIT) Then Exit Do
End If
PeekMessage checks the queue for a message if there is one then it returns non zero and places the message info in the first parameter (message) the next 3 zero are arbitrary and the final parameter is to choose to remove or not remove the message form the queue
TranslateMessage is only needed if there is a menu but it doesn't slow down the loop so its good to just leave it
DispatchMessage sends the app the message received with PeekMessage so the program can handle it
The if statement at the end is needed if quit is called then this will exit the main loop other wise the program will continue to run
If PeekMessage(Message, 0, 0, 0, PM_NOREMOVE) Then
  DoEvents
End If
DoEvents removes and dispatches all the messages in the queue while DispatchMessage will only dispatch the last message received which slows down the loop especially if the loop takes a while to make one full loop around because there will probably be more event messages in the queue
Inputting this also means that there has to be a condition for the do loop to quit (probably in form_unload) to exit the loop because if there was a check for a quit message it could be under another message and doevents would remove it from the queue and the loop would become endless
If PeekMessage(Message, 0, 0, 0, PM_NOREMOVE) Then
  If GetMessage(Message, 0, 0, 0) = 0 Then Exit Do
  DispatchMessage Message
  If Message.Message = MY_WM_QUIT Then
   Exit Do
  End If
Else
End If
This is almost the same as the first loop except that it has the extra error check and the main part of the program goes after the else statement its slightly slower than the first I would opt to use the first instead of this unless error checking is your main concern
I also put this here to show there are many ways to get/manage event messages
MY INSPIRATION
If GetQueueStatus(QS_ALLEVENTS) Then
  DoEvents
End If
GetQueueStatus(qs_event) checks the queue for a message of the users choice although this is faster than just doing a standard DoEvents the GetQueueStatus approach is slower than the first 3 examples
Using GetQueueStatus before DoEvents was posed earlier by John Galanopoulos and a class was made on it by JohnB which made me think about an even faster way of doing this or a way to eliminate DoEvents all together I have done some research and my first and third examples are how messages should be handled by games made in C++
Please don't vote for this code it was just posted to be shared with everyone and not originally thought up by me(mostly just translated from C++ and John G's excellent article) so I don't feel I really deserve any credit for this but you can vote for my previous entry warbotz if you like it:)
```

