' Goal Seek example

Sub goalSeek()
	Range("H13").GoalSeek Goal:=Range("I8"), ChangingCell:=Range("H10")
End Sub 