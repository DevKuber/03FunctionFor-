Attribute VB_Name = "Module1"
Option Explicit



Public Function RemoveLeftRight(Ref As String) As String
Dim StrLen As Integer
StrLength = Len(Ref)
i As Integer
For i = 1 To StrLength
                
                If Mid(Ref, i, 1) = "<" Then
                 
                                           
                    Mid(Ref, i, 1) = ""
                    Mid(Ref, i + 1, 1) = ""
                    Mid(Ref, i + 2, 1) = ""
                    Mid(Ref, i + 3, 1) = ","
                    
                Next i
                
                
                RemoveLeftRight = Ref
  
                    
End Function



