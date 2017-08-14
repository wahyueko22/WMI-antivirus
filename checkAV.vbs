Dim objWMIService, objItem, colItems

Set objWMIService = GetObject("winmgmts:\\.\root\SecurityCenter2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM AntivirusProduct")

Dim objFSO, outputFile, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
outputFile = "C:\ProgramData\AVHex.out"

Set objFile = objFSO.CreateTextFile(outputFile, true)

For Each objItem in colItems
  Dim displayName, productState, stateStatus, stateHex, stateUpToDate, stateEnabled
  displayName = objItem.displayName
  productState = objItem.productState

  stateHex = Hex(productState)

  If (Len(stateHex) = 5) Then
    stateHex = "0" & stateHex
  End If
  Wscript.Echo displayName & ": " & productState & " [Hex: " & stateHex & "]"

  If (Right(stateHex, 2) = 00) Then
    stateUpToDate = "Definition up to date"
  Else
    stateUpToDate = "Definition not up to date"
  End If

  If (Mid(stateHex, 3, 1) = 1) Then
    stateEnabled = "Protection enabled"
  ElseIf (Mid(stateHex, 3, 1) < 1) Then
    stateEnabled = "Protection disabled"
  ElseIf (Mid(stateHex, 3, 1) = 2) Then
    stateEnabled = "Protection disabled"
  Else
    stateEnabled = "Not all protection enabled"
  End If

  stateStatus = "displayName: " & displayName & ", productState: " & productState & ", status: " & stateEnabled & " and " & stateUpToDate
  
  Wscript.Echo stateStatus

  objFile.Write stateStatus & vbCrLf

Next

objFile.Close