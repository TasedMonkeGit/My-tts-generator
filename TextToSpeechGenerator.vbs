Dim message, sapi
message = InputBox("Type something for the computer to say:", "Friendly TTS")

If message <> "" Then
    Set sapi = CreateObject("SAPI.SpVoice")
    sapi.Rate = -1
    sapi.Volume = 80
    sapi.Speak message
End If
