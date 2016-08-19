' Homestead / Vagrant - starter 
' This is a helper that allows you to start vagrant / homestead without a cmd or powershell
' 
' Requires Homestead and vagrant be installed 
' Note: only works with virtualbox
' Author Luuk Verhoeven 
' Date 2016
' Version 1.0

Set objShell = CreateObject("WScript.Shell")
Set objFSO=CreateObject("Scripting.FileSystemObject")

Dim strMsg,inp01,strTitle,strFlag

strTitle = "Vagrant OPTIONS"

strMsg = "Enter S to start" & vbCr
strMsg = strMsg & "Enter T to Terminal ssh" & vbCR
strMsg = strMsg & "Enter D to Destroy and Start" & vbCR
strMsg = strMsg & "Enter P to reload provision" & vbCR
strMsg = strMsg & "Enter X to suspend" & vbCR
strFlag = False

' Get use profile
userProfilePath = objShell.ExpandEnvironmentStrings("%UserProfile%")

outFile= userProfilePath + "\tmp.bat"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "@echo off" & vbCrLf
objFile.Write "cd " & userProfilePath & "\Homestead" & vbCrLf

Do While strFlag = False

inp01 = InputBox(strMsg,"Make your selection")

Select Case inp01
    Case "t"
      	objFile.Write "echo vagrant ssh" & vbCrLf
      	objFile.Write "vagrant ssh" & vbCrLf
        strFlag = True
    Case "s"
    	objFile.Write "vagrant version"  & vbCrLf
      	objFile.Write "echo vagrant up" & vbCrLf
      	objFile.Write "vagrant up" & vbCrLf
      	objFile.Write "vagrant ssh" & vbCrLf
        strFlag = True
    Case "d"
     	objFile.Write "echo vagrant destroy -f" & vbCrLf
        objFile.Write "vagrant destroy -f" & vbCrLf
        objFile.Write "pause" & vbCrLf
        strFlag = True
    Case "x"
    	objFile.Write "echo vagrant suspend" & vbCrLf
        objFile.Write "vagrant suspend" & vbCrLf
        strFlag = True
    Case "p"
    	objFile.Write "echo vagrant reload --provision" & vbCrLf
        objFile.Write "vagrant reload --provision" & vbCrLf
        objFile.Write "vagrant ssh" & vbCrLf
        strFlag = True
    Case Else
        MsgBox "You made an incorrect selection!",64,strTitle
        strFlag = True
End Select
Loop

' Closing

objFile.Close

objShell.Run outFile

Wscript.Quit