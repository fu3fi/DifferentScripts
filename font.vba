Sub RandomFont()
  Application.ScreenUpdating = False

  Set objDoc = ActiveDocument
  Set objRandom = CreateObject("System.Random")

  Set objRange = objDoc.Range()
  Set colCharacters = objRange.Characters
  
  For Each strCharacter In colCharacters
      'strCharacter.Font.Reset
      strCharacter.Font.Scaling = 100 + objRandom.Next_2(-50, 50) / 8
      strCharacter.Font.Position = objRandom.Next_2(-200, 300) / 700
      strCharacter.Font.Size = strCharacter.Font.Size + objRandom.Next_2(-300, 400) / 400
      strCharacter.Font.Kerning = 12 + objRandom.Next_2(-10, 40) / 5
      Select Case objRandom.Next_2(1, 5)
        Case 1
          strCharacter.Font.Name = "ZimM-1"
        Case 2
          strCharacter.Font.Name = "ZimM-2"
        Case 3
          strCharacter.Font.Name = "ZimM-3"
        Case 4
          strCharacter.Font.Name = "ZimM-4"
      End Select
  Next
  
  Application.ScreenUpdating = True
End Sub