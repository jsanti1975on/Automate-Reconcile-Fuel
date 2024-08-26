Option Explicit

Sub Pause(duration_ms As Double)
    Dim start_time As Double
    start_time = Timer
    Do
        DoEvents
    Loop Until (Timer - start_time) * 1000 >= duration_ms
End Sub

Sub SimulateConversation()
    Dim bubble1 As Shape, bubble2 As Shape
    Dim i As Integer
    Dim messages1 As Variant
    Dim messages2 As Variant
    
    ' Set the text bubble shapes
    Set bubble1 = ActiveSheet.Shapes("TextBubble1")
    Set bubble2 = ActiveSheet.Shapes("TextBubble2")
    
    ' Set font size for the text bubbles
    bubble1.TextFrame2.TextRange.Font.Size = 20 ' Adjust font size as needed
    bubble2.TextFrame2.TextRange.Font.Size = 20 ' Adjust font size as needed
    
    ' Define the messages for the conversation
    messages1 = Array("Hello, my name is Logger!", "Now let's use math to assist with closing", "Paste over today's fuel numbers", "After the paste process!", "If you see discrepancies over .20", "If need be, we can troubleshoot further on sheet2")
    messages2 = Array("I am here to help you with troubleshooting!", "Let's go ahead and start with today's data", "I have been blinking the fields RED, to show where data goes.", "Select the green button over my head!", "We may have completed the process!", "Sheet 2 has some additional help for further troubleshooting.")
    
    ' Loop to animate the conversation and blink the range
    For i = LBound(messages1) To UBound(messages1)
        ' Start blinking the range
        BlinkRange "R2:V2", 5000 ' Blink for 5000 milliseconds (5 seconds)
        
        ' Show the first text bubble with specific message
        bubble1.TextFrame2.TextRange.Text = "LOGGER: " & messages1(i)
        bubble1.Visible = msoTrue
        AnimateSmileyFace 8000 ' Animate the smiley face while displaying the message
        bubble1.Visible = msoFalse
        
        ' Show the second text bubble with specific message
        bubble2.TextFrame2.TextRange.Text = "LOGGER: " & messages2(i)
        bubble2.Visible = msoTrue
        AnimateSmileyFace 7000 ' Animate the smiley face while displaying the message
        bubble2.Visible = msoFalse
        
        ' Stop blinking the range
        BlinkRange "R2:V2", 0 ' Set to 0 to stop blinking
    Next i
End Sub

Sub BlinkRange(rangeAddress As String, blinkDuration As Double)
    Dim cell As Range
    Dim endTime As Double
    Dim blinkOn As Boolean
    
    ' Get the range
    Set cell = ActiveSheet.Range(rangeAddress)
    
    ' Set the end time
    endTime = Timer + (blinkDuration / 1000)
    
    ' Blink the range
    Do While Timer < endTime And blinkDuration > 0
        ' Toggle the blink state
        blinkOn = Not blinkOn
        If blinkOn Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color
        Else
            cell.Interior.ColorIndex = xlNone ' No fill
        End If
        ' Pause for 200 milliseconds
        Pause 200
    Loop
    
    ' Ensure the range is not filled at the end
    cell.Interior.ColorIndex = xlNone
End Sub

Sub AnimateSmileyFace(duration_ms As Double)
    Dim mouthOpen As Shape
    Dim mouthClosed As Shape
    Dim startTime As Double
    
    ' Set the shapes
    Set mouthOpen = ActiveSheet.Shapes("MouthOpen")
    Set mouthClosed = ActiveSheet.Shapes("MouthClosed")
    
    ' Initial visibility
    mouthOpen.Visible = msoTrue
    mouthClosed.Visible = msoFalse
    
    ' Get the start time
    startTime = Timer
    
    ' Loop to animate the mouth opening and closing
    Do While (Timer - startTime) * 1000 < duration_ms
        ' Toggle visibility
        mouthOpen.Visible = Not mouthOpen.Visible
        mouthClosed.Visible = Not mouthClosed.Visible
        
        ' Pause for a short duration (500 milliseconds)
        Pause 500
    Loop
    
    ' Ensure mouth is closed at the end
    mouthOpen.Visible = msoFalse
    mouthClosed.Visible = msoTrue
End Sub

Sub SimulateConversationWithAnimation()
    ' Start the conversation with animation
    Call SimulateConversation
End Sub

