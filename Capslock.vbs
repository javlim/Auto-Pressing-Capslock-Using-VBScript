Dim toggleCount, ie, terminateScript

toggleCount = 0
terminateScript = False

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True

ie.Navigate "about:blank"
Do While ie.Busy
    WScript.Sleep 100
Loop

Set doc = ie.Document
doc.Write "<html><body><h1 id='Counter'>Number of Times Capslock is Pressed: 0 </h1><button onclick='window.close();'>Stop</button></body></html>"
doc.close

Set counter = doc.getElementById("Counter")

Do
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{CAPSLOCK}"
    toggleCount = toggleCount+1

    Counter.InnerText = "Number of Times Capslock is Pressed: " & toggleCount

    For i = 1 to 2400
        WScript.Sleep 100

        on Error Resume Next
        If ie Is Nothing Or ie.Document Is Nothing Then
            terminateScript = True
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
    Next

Loop Until terminateScript

ie.Quit