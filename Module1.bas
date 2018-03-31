Attribute VB_Name = "Module1"
Option Explicit

'//===============================================
'//This function create Random number in special range
'//Count     ==> count of number that must created
'//Min       ==> Minimume of number that can be created
'//Max       ==> Maximume of number that can be created
'//Result()  ==> A byref array for put result in it and return to user



Public Function Random_X(ByVal Count As Long, ByVal Min As Long, ByVal Max As Long, ByRef Result() As Long, ByVal Sort_Array As Boolean) As Boolean

Dim i As Long
Dim Top_Array As Long
Dim Rand_Num As Long

    Randomize  '//Randomize Timer

    '//============================
    '//First check that count in range (MAX-MIN)
    
        If Count > (Max - Min) Then
        
            Random_X = False
            Exit Function
            
        Else
        
            Random_X = True
        
        End If
    
    '//============================

    Top_Array = 0

    ReDim Result(Count - 1) '//Redim Empty Array and Fit it to Count
    
    For i = LBound(Result) To UBound(Result)
    
Repeat:

        Rand_Num = Rnd() * Max
        Rand_Num = Rand_Num + Max '//Go Number larger than max
        
        Do While (Rand_Num < Min Or Rand_Num > Max)
        
            Rand_Num = Rand_Num - (Max - Min) '// IF Rand number is out of range , come it in range
        
        Loop
         
        If In_Array_X(Result, Rand_Num, i) = False Then '//IF Not exist then push it into array
         
            Result(i) = Rand_Num
                
        Else
        
            GoTo Repeat
         
        End If
    
    Next
    
    If Sort_Array = True Then Sort Result         '//If Sort =True then Sort result array

End Function

'//=======================================
'//This function get a byref array and a num
'//Check the num exist in array

Public Function In_Array_X(ByRef Arr_Name() As Long, ByVal num As Long, ByVal Top_Arr As Long) As Boolean

Dim i As Long

    In_Array_X = False
    
    If Top_Arr > UBound(Arr_Name) Then Top_Arr = UBound(Arr_Name)

    For i = LBound(Arr_Name) To Top_Arr
    
         If Arr_Name(i) = num Then
         
            In_Array_X = True
            Exit For
         End If
      Next i
End Function


'//=======================================
'//This Function get a byref array and sort it

Public Sub Sort(ByRef Sort_Arr() As Long)

Dim i As Long, j As Long
Dim Temp As Long

    For i = UBound(Sort_Arr) - 1 To LBound(Sort_Arr) Step -1
    
         For j = 0 To i Step 1
         
                If Sort_Arr(j) > Sort_Arr(j + 1) Then
                
                    Temp = Sort_Arr(j)
                    Sort_Arr(j) = Sort_Arr(j + 1)
                    Sort_Arr(j + 1) = Temp
                
                End If
                
         Next
    
    Next

End Sub



