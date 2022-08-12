
'By listing the first six prime numbers: 2, 3, 5, 7, 11, and 13, we can see that the 6th prime is 13.
'What is the 10001st prime number?


iNumberToLookAt = 1
iPrimesFound = 0
iMaxPrimeSeq = 10001

Do Until iPrimesFound >= iMaxPrimeSeq
	iNumberToLookAt = iNumberToLookAt + 1
	If IsPrime(iNumberToLookAt) = True Then
		iPrimesFound = iPrimesFound + 1
	End If
Loop

MsgBox "Done!" & chr(13) & iNumberToLookAt & chr(13) & iPrimesFound


'-----------------------------
Function IsPrime(iNumberToLookAt)
	iGoUpToThisNumber = Int(iNumberToLookAt/2)
	bResult = True
	For i = 2 To iGoUpToThisNumber
		If iNumberToLookAt Mod i = 0 Then
			bResult = False
			Exit For
		End If
	Next
	IsPrime = bResult
End Function
