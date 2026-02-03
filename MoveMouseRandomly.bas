Attribute VB_Name = "Module14"
Declare PtrSafe Sub SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long)
Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Const VK_ESCAPE = &H1B

Sub MoveMouseRandomly()
Attribute MoveMouseRandomly.VB_ProcData.VB_Invoke_Func = "W\n14"
    Dim centerX As Long, centerY As Long
    Dim baseRadius As Long
    Dim angle As Double
    Dim x As Long, y As Long
    Dim i As Long
    Dim randRadius As Double, randAngleOffset As Double
    Dim screenWidth As Long, screenHeight As Long
    Dim t As Double
    Dim amplitudeX As Double, amplitudeY As Double
    Dim frequencyX As Double, frequencyY As Double

    centerX = 500
    centerY = 500
    baseRadius = 100
    i = 0
    
     ' Set screen dimensions (adjust if needed)
    screenWidth = 1920
    screenHeight = 1080

    ' Loop parameters
    amplitudeX = screenWidth / 2
    amplitudeY = screenHeight / 2
    frequencyX = 0.02
    frequencyY = 0.03
    t = 0
    

    Do While True
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then Exit Sub

        ' circle with shake
        'randRadius = baseRadius + Rnd() * 30 - 15
        'randAngleOffset = Rnd() * 10 - 5
        'angle = (i + randAngleOffset) * 3.14159 / 180
        'x = centerX + randRadius * Cos(angle)
        'y = centerY + randRadius * Sin(angle)
        'SetCursorPos x, y
        'Delay 0.05
        
        'circle
        'angle = i * 3.14159 / 180
        'x = centerX + radius * Cos(angle)
        'y = centerY + radius * Sin(angle)
        'SetCursorPos x, y
        'Delay 0.05

        'i = i + 5
        'If i >= 360 Then i = 0
        
          ' Generate looping motion
        x = screenWidth / 2 + amplitudeX * Sin(frequencyX * t)
        y = screenHeight / 2 + amplitudeY * Cos(frequencyY * t)

        SetCursorPos x, y
        Delay 0.05

        t = t + 1
        
    Loop
End Sub

Sub Delay(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

