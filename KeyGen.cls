VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'SecureKey (C) 2003 - P-Soft
'Coded by Pio; ydj@aol.com; http://p-soft.shockrock.net

Private Function SubKeys(lngSeed As Long) As String
Dim i As Long, r As Long, Keys(1 To 4) As Variant, _
strCurKey As String

    Select Case lngSeed
        Case 1
            Keys(1) = Array("0", "3", "8", "6", "1")
            Keys(2) = Array("6", "2", "4", "0", "3")
            Keys(3) = Array("E", "X", "C", "M", "J")
            Keys(4) = Array("6", "2", "4", "0", "3")
        Case 2
            Keys(1) = Array("0", "1", "6", "5", "8")
            Keys(2) = Array("9", "1", "0", "5", "7")
            Keys(3) = Array("R", "Z", "K", "F", "M")
            Keys(4) = Array("8", "1", "3", "0", "7")
        Case 3
            Keys(1) = Array("8", "1", "4", "2", "3")
            Keys(2) = Array("1", "3", "0", "0", "4")
            Keys(3) = Array("P", "S", "Q", "K", "D")
            Keys(4) = Array("0", "6", "3", "0", "1")
        Case 4
            Keys(1) = Array("3", "0", "5", "2", "4")
            Keys(2) = Array("0", "0", "1", "9", "3")
            Keys(3) = Array("R", "L", "H", "K", "Z")
            Keys(4) = Array("0", "9", "0", "0", "6")
    End Select
    
    
    For i = 1 To 4
    DoEvents
        r = RandomNumber(5)
            If r = 0 Then
                r = 1
            End If
        
        strCurKey = strCurKey & Keys(i)(r)

    Next
    
    SubKeys = strCurKey
    

End Function
Public Function ValidKey(strKey As String) As Boolean 'Validate Key
Dim Keys(1 To 4) As Variant, i As Long, _
strBuffer As String, k As Long, blnValid As Boolean



ValidKey = True
    
    If InStr(strKey, "-") Then 'Make sure the key has a -, else it's invalid
        DoEvents
    Else
        ValidKey = False
        Exit Function
    End If
    
    
    
    If UBound(Split(strKey, "-")) + 1 <> 4 Then 'make sure that there are only 4 SubKeys, else it's invalid
        ValidKey = False
    Exit Function
    End If
    
    
    
    For i = 1 To 4 'Make the Key into SubKeys
    Keys(i) = Split(strKey, "-")(i - 1)
    Next
    
    
    For k = 1 To 4 'Go through each SubKey
    
    blnValid = False 'Reset the variable
    
        For i = 1 To 4 'Go through each letter in the SubKey
        DoEvents
            blnValid = False 'Reset the variable
            strBuffer = Mid(Keys(k), i, 1)
            
            If Len(Keys(k)) > 4 Then
            blnValid = False
            ValidKey = False
            Exit Function
            End If
            
            
            Select Case k
                    
                Case 1
                    
                    Select Case i
                    
                        Case 1
                        
                            'Keys(1) = Array("0", "3", "8", "6", "1")
                            
                            If strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            ElseIf strBuffer = "8" Then
                                blnValid = True
                            ElseIf strBuffer = "6" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 2
                        
                            'Keys(2) = Array("6", "2", "4", "0", "3")
                            
                            If strBuffer = "6" Then
                                blnValid = True
                            ElseIf strBuffer = "2" Then
                                blnValid = True
                            ElseIf strBuffer = "4" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 3
                        
                            'Keys(3) = Array("E", "X", "C", "M", "J")
                            
                            If strBuffer = "E" Then
                                blnValid = True
                            ElseIf strBuffer = "X" Then
                                blnValid = True
                            ElseIf strBuffer = "C" Then
                                blnValid = True
                            ElseIf strBuffer = "M" Then
                                blnValid = True
                            ElseIf strBuffer = "J" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 4
                        
                            'Keys(4) = Array("6", "2", "4", "0", "3")
                            
                            If strBuffer = "6" Then
                                blnValid = True
                            ElseIf strBuffer = "2" Then
                                blnValid = True
                            ElseIf strBuffer = "4" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                    End Select
                    
                Case 2
                    
                    Select Case i

                        Case 1
                        
                            'Keys(1) = Array("0", "1", "6", "5", "8")
                            
                            If strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "6" Then
                                blnValid = True
                            ElseIf strBuffer = "5" Then
                                blnValid = True
                            ElseIf strBuffer = "8" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 2
                        
                            'Keys(2) = Array("9", "1", "0", "5", "7")
                            
                            If strBuffer = "9" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "5" Then
                                blnValid = True
                            ElseIf strBuffer = "7" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 3
                        
                            'Keys(3) = Array("R", "Z", "K", "F", "M")
                            
                            If strBuffer = "R" Then
                                blnValid = True
                            ElseIf strBuffer = "Z" Then
                                blnValid = True
                            ElseIf strBuffer = "K" Then
                                blnValid = True
                            ElseIf strBuffer = "F" Then
                                blnValid = True
                            ElseIf strBuffer = "M" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 4
                        
                            'Keys(4) = Array("8", "1", "3", "0", "7")
                            
                            If strBuffer = "8" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "7" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                    End Select

                
                Case 3
                    
                    Select Case i
                        
                        Case 1
                        
                            'Keys(1) = Array("8", "1", "4", "2", "3")
                            
                            If strBuffer = "8" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "4" Then
                                blnValid = True
                            ElseIf strBuffer = "2" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 2
                        
                            'Keys(2) = Array("1", "3", "0", "0", "4")
                            
                            If strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "4" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 3
                        
                            'Keys(3) = Array("P", "S", "Q", "K", "D")
                            
                            If strBuffer = "P" Then
                                blnValid = True
                            ElseIf strBuffer = "S" Then
                                blnValid = True
                            ElseIf strBuffer = "Q" Then
                                blnValid = True
                            ElseIf strBuffer = "K" Then
                                blnValid = True
                            ElseIf strBuffer = "D" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 4
                        
                            'Keys(4) = Array("0", "6", "3", "0", "1")
                            
                            If strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "6" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                    End Select
                
                Case 4
                    
                    Select Case i

                        Case 1
                        
                            'Keys(1) = Array("3", "0", "5", "2", "4")
                            
                            If strBuffer = "3" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "5" Then
                                blnValid = True
                            ElseIf strBuffer = "2" Then
                                blnValid = True
                            ElseIf strBuffer = "4" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 2
                            
                            'Keys(2) = Array("0", "0", "1", "9", "3")
                            
                            If strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "1" Then
                                blnValid = True
                            ElseIf strBuffer = "9" Then
                                blnValid = True
                            ElseIf strBuffer = "3" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 3
                            
                            'Keys(3) = Array("R", "L", "H", "K", "Z")
                            
                            If strBuffer = "R" Then
                                blnValid = True
                            ElseIf strBuffer = "L" Then
                                blnValid = True
                            ElseIf strBuffer = "H" Then
                                blnValid = True
                            ElseIf strBuffer = "K" Then
                                blnValid = True
                            ElseIf strBuffer = "Z" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                        Case 4
                            
                            'Keys(4) = Array("0", "9", "0", "0", "6")
                            
                            If strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "9" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "0" Then
                                blnValid = True
                            ElseIf strBuffer = "6" Then
                                blnValid = True
                            End If
                                
                                If blnValid = False Then
                                    ValidKey = False
                                    Exit Function
                                End If
                                
                    End Select
                
            End Select
            
        Next
        
    Next
    
    

End Function
Public Function MakeKey() As String 'Create a Key
    MakeKey = SubKeys(1) & "-" & SubKeys(2) & "-" & SubKeys(3) & "-" & SubKeys(4)
End Function
Private Function RandomNumber(Max As Variant) As Long
Dim RndNum As Variant
    RndNum = Int((Val(Max) * Rnd))
        RandomNumber = CLng(RndNum)
End Function
