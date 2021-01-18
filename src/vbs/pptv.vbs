On Error Resume Next
Set ppt = WScript.CreateObject("PowerPoint.Application")

if Err <> 0 Then
    WSH.Echo Err.Description
else
    WSH.Echo ppt.Version
End If