
'The prime factors of 13195 are 5, 7, 13 and 29.
'What is the largest prime factor of the number 600851475143 ?

iLargestPrimeFactor = 0
'iMegaNumber = 13195
iMegaNumber = cDbl(600851475143)
iMegaNumber = cDbl(8462696833)
iMegaNumber = cDbl(408464633)
'iMegaNumber = cDbl(10086647)
iMinPrimeFactor = 1471
iMinPrimeFactor = 839
'6857

iWorkingNumber = cDbl(iMegaNumber)
iLowerFloor = 73

iSqrRootOfMega = cDbl(Sqr(iMegaNumber))
iSqrRootOfMega = iMegaNumber

For iNumberToLookAt = iLowerFloor To iSqrRootOfMega
	
	If IsPrime(iNumberToLookAt) = True Then
		iModulus = iWorkingNumber Mod iNumberToLookAt
		If iWorkingNumber = iNumberToLookAt Then
			iLargestPrimeFactor = iNumberToLookAt
			Exit For
		ElseIf iModulus = 0 Then
			iWorkingNumber = iWorkingNumber / iNumberToLookAt
		End If
	End If

Next

'sBleck = chr(13) & "------" & chr(13) & iMegaNumber & chr(13) & iNumberToLookAt & chr(13) & iWorkingNumber & chr(13) & iCount & chr(13) & iModulus

MsgBox "Done!" & chr(13) & iLargestPrimeFactor


'-----------------------------
Function IsPrime(iNumberToLookAt)
	iGoUpToThisNumber = Int(Sqr(iNumberToLookAt))
	bResult = True
	For i = 2 To iGoUpToThisNumber
		If iNumberToLookAt Mod i = 0 Then
			bResult = False
			Exit For
		End If
	Next
	IsPrime = bResult
End Function

