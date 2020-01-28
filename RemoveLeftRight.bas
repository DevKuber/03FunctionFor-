Attribute VB_Name = "Module1"
Option Explicit



Public Function RemoveLeftRight(Word As String) As String
Dim StrLen As Integer
StrLength = Len(Word)
i As Integer
For i = 1 To StrLength
                
                If Mid(Word, i, 1) = "<" Then                       'Elenxos an Ksekinaei me <'
                        
                        
                                           
                    Mid(Word, i, 1) = " "
                    Mid(Word, i + 1, 1) = " "
                    Mid(Word, i + 2, 1) = " "
                    If Mid(Word, i + 1, 1) = "/" Then               'Elenxos an exei /'
                    Mid(Word, i + 3, 1) = ","
                    
                Next i
    StrLength = Len(Word)
                
                Mid(Word, StrLength, 1) = " "                        'Delete ton Teleuteou Comma'
                
                RemoveLeftRight = Word
  
                    
End Function



