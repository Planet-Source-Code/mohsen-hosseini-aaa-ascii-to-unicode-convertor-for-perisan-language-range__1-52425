Attribute VB_Name = "Cnvrt2Uni"
Function CUni(str1 As String) As String
Dim Char, struni As String
Dim ch1 As Integer
For i = 1 To Len(str1)
    Char = Left(str1, 1)
    
    Select Case Char
        'Case " ": ch1 = &H20
        Case "H": ch1 = &H622
        Case "h": ch1 = &H627
        Case "f", "F": ch1 = &H628
        Case "`": ch1 = &H67E
        Case "j", "J": ch1 = &H62A
        Case "e", "E": ch1 = &H62B
        Case "[": ch1 = &H62C
        Case "]": ch1 = &H6286
        Case "p", "P": ch1 = &H62D
        Case "o", "O": ch1 = &H62E
        Case "n", "N": ch1 = &H62F
        Case "b", "B": ch1 = &H630
        Case "v", "V": ch1 = &H631
        Case "c", "C": ch1 = &H632
        Case "\": ch1 = &H698
        Case "s", "S": ch1 = &H633
        Case "a", "A": ch1 = &H634
        Case "w", "W": ch1 = &H635
        Case "q", "Q": ch1 = &H636
        Case "x", "X": ch1 = &H637
        Case "z", "Z": ch1 = &H638
        Case "u", "U": ch1 = &H639
        Case "y", "Y": ch1 = &H63A
        Case "t", "T": ch1 = &H641
        Case "r", "R": ch1 = &H642
        Case ";": ch1 = &H6A9
        Case "'": ch1 = &H6AF
        Case "g", "G": ch1 = &H644
        Case "l", "L": ch1 = &H645
        Case "k", "K": ch1 = &H646
        Case ",": ch1 = &H648
        Case ">": ch1 = &H623
        Case "<": ch1 = &H624
        Case "i", "I": ch1 = &H647
        Case "d", "D": ch1 = &H6CC
        Case "M": ch1 = &H626
        Case "m": ch1 = &H621
        Case "0": ch1 = &H6F0
        Case "1": ch1 = &H6F1
        Case "2": ch1 = &H6F2
        Case "3": ch1 = &H6F3
        Case "4": ch1 = &H6F4
        Case "5": ch1 = &H6F5
        Case "6": ch1 = &H6F6
        Case "7": ch1 = &H6F7
        Case "8": ch1 = &H6F8
        Case "9": ch1 = &H6F9
        Case Else: ch1 = Asc(Char)
        
End Select
    
    'List1.AddItem ChrW(ch1)
    str1 = Right(str1, Len(str1) - 1)
    CUni = CUni + ChrW(ch1)
    
Next i
End Function
