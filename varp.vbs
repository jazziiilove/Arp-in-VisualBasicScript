Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("c:\arp.txt", ForReading)

Dim arrFileLines()
i = 0
Do Until objFile.AtEndOfStream
ReDim Preserve arrFileLines(i)
arrFileLines(i) = objFile.ReadLine
i = i + 1
Loop
objFile.Close

Set objShell = CreateObject("Wscript.Shell") objShell.Run ("arp -s " & arrFileLines(0))