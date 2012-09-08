' iTunes playcount editor
'
' todo - work out why it's sometimes failing

Dim iTunesApp, selectedTracks, newPlayCount, numEdits
Set iTunesApp = WScript.CreateObject("iTunes.Application")
Set selectedTracks = iTunesApp.SelectedTracks

intHighNumber = 90
intLowNumber = 30
numEdits = 0

For Each IITTrack In selectedTracks
	numEdits = numEdits+1
	Randomize
	newPlayCount = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)

	If Len(newPlayCount) > 0 Then
		If IsNumeric(newPlayCount) Then
			If newPlayCount >= 0 Then 
				IITTrack.PlayedCount = newPlayCount
			End If
		End If
	Else
		Exit For
	End If
	
	'WScript.Echo "Process Complete"
Next
WScript.Echo "Number of Edits: " & numEdits