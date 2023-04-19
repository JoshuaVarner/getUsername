' GetUsername.vbs
Option Explicit
Dim objNetwork, strUsername

' Create the WScript.Network object
Set objNetwork = CreateObject("WScript.Network")

' Get the username of the current logged-in user
strUsername = objNetwork.UserName

' Display the username in a message box
MsgBox "The current logged-in user is: " & strUsername, vbInformation, "Current User"

' Clean up
Set objNetwork = Nothing
