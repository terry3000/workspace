Sub 对比数据()
    '
    ' 对比数据
    '
    Range("A2:AA33").Interior.Pattern = xlPatternNone
    Range("H2:H33,L2:L33,P2:P33,T2:T33,X2:X33").Clear
    Columns("H").ColumnWidth = 1
    Columns("L").ColumnWidth = 1
    Columns("P").ColumnWidth = 1
    Columns("T").ColumnWidth = 1
    Columns("X").ColumnWidth = 1
    Range("A12:AA12").Interior.Color = RGB(192, 192, 192)
    Range("A23:AA23").Interior.Color = RGB(192, 192, 192)
    Range("A34:AA34").Interior.Color = RGB(128, 138, 135)
    '版本1  相同位置互为相反数
    shuzu1("E2") : shuzu1("F2") : shuzu1("G2") : shuzu2("E3") : shuzu2("F3") : shuzu2("G3") : shuzu3("E4") : shuzu3("F4") : shuzu3("G4")
    '版本20   找到93
    banben2001("E2") : banben2001("I2") : banben2001("M2") : banben2001("Q2") : banben2001("U2") : banben2001("Y2")
    banben2001("E6") : banben2001("I6") : banben2001("M6") : banben2001("Q6") : banben2001("U6") : banben2001("Y6")
    banben2001("E9") : banben2001("I9") : banben2001("M9") : banben2001("Q9") : banben2001("U9") : banben2001("Y9")
    banben2001("E13") : banben2001("I13") : banben2001("M13") : banben2001("Q13") : banben2001("U13") : banben2001("Y13")
    banben2001("E17") : banben2001("I17") : banben2001("M17") : banben2001("Q17") : banben2001("U17") : banben2001("Y17")
    banben2001("E20") : banben2001("I20") : banben2001("M20") : banben2001("Q20") : banben2001("U20") : banben2001("Y20")
    banben2001("E24") : banben2001("I24") : banben2001("M24") : banben2001("Q24") : banben2001("U24") : banben2001("Y24")
    banben2001("E28") : banben2001("I28") : banben2001("M28") : banben2001("Q28") : banben2001("U28") : banben2001("Y28")
    banben2001("E31") : banben2001("I31") : banben2001("M31") : banben2001("Q31") : banben2001("U31") : banben2001("Y31")
    '版本20   找到31
    banben2002("E2") : banben2002("I2") : banben2002("M2") : banben2002("Q2") : banben2002("U2") : banben2002("Y2")
    banben2002("E6") : banben2002("I6") : banben2002("M6") : banben2002("Q6") : banben2002("U6") : banben2002("Y6")
    banben2002("E9") : banben2002("I9") : banben2002("M9") : banben2002("Q9") : banben2002("U9") : banben2002("Y9")
    banben2002("E13") : banben2002("I13") : banben2002("M13") : banben2002("Q13") : banben2002("U13") : banben2002("Y13")
    banben2002("E17") : banben2002("I17") : banben2002("M17") : banben2002("Q17") : banben2002("U17") : banben2002("Y17")
    banben2002("E20") : banben2002("I20") : banben2002("M20") : banben2002("Q20") : banben2002("U20") : banben2002("Y20")
    banben2002("E24") : banben2002("I24") : banben2002("M24") : banben2002("Q24") : banben2002("U24") : banben2002("Y24")
    banben2002("E28") : banben2002("I28") : banben2002("M28") : banben2002("Q28") : banben2002("U28") : banben2002("Y28")
    banben2002("E31") : banben2002("I31") : banben2002("M31") : banben2002("Q31") : banben2002("U31") : banben2002("Y31")
    '版本20   找到“红零”
    banben2003("E2") : banben2003("I2") : banben2003("M2") : banben2003("Q2") : banben2003("U2") : banben2003("Y2")
    banben2003("E6") : banben2003("I6") : banben2003("M6") : banben2003("Q6") : banben2003("U6") : banben2003("Y6")
    banben2003("E9") : banben2003("I9") : banben2003("M9") : banben2003("Q9") : banben2003("U9") : banben2003("Y9")
    banben2003("E13") : banben2003("I13") : banben2003("M13") : banben2003("Q13") : banben2003("U13") : banben2003("Y13")
    banben2003("E17") : banben2003("I17") : banben2003("M17") : banben2003("Q17") : banben2003("U17") : banben2003("Y17")
    banben2003("E20") : banben2003("I20") : banben2003("M20") : banben2003("Q20") : banben2003("U20") : banben2003("Y20")
    banben2003("E24") : banben2003("I24") : banben2003("M24") : banben2003("Q24") : banben2003("U24") : banben2003("Y24")
    banben2003("E28") : banben2003("I28") : banben2003("M28") : banben2003("Q28") : banben2003("U28") : banben2003("Y28")
    banben2003("E31") : banben2003("I31") : banben2003("M31") : banben2003("Q31") : banben2003("U31") : banben2003("Y31")
    '版本2  E7行 绝对值上下相等
    ' shangxia ("E7"):' shangxia("F7"): ' shangxia("G7")
    '
    '版本3 版本4
    'E7行 互为相反数  且 上下均为互为相反数
    'E7行 互为相反数  且 上下颠倒
    '   shangxia7 ("E7"):'   shangxia7("F7"): '   shangxia7("G7")
    '版本10-1  A1 B2 C3 绝对值相等
    '  banben101 ("E9"):  banben101 ("F9"):  banben101 ("G9"): banben101 ("E2"):  banben101 ("F2"):'  banben101("G2")
    ' '版本10-2  A1 B3 C2  绝对值相等
    '  banben102("E9"):  banben102("F9"):  banben102("G9"):  banben102("E2"):  banben102("F2"):  banben102("G2")
    '  '版本10-3  A2 B1 C3
    '  banben103("E9"):  banben103("F9"):  banben103("G9"):  banben103("E2"):  banben103("F2"):  banben103("G2")
    '  '版本10-4  A2 B3 C1
    '  banben104("E9"):  banben104("F9"):  banben104("G9"):  banben104("E2"):  banben104("F2"):  banben104("G2")
    '  版本10-5  A3 B1  C2
    '  banben105("E9"):  banben105("F9"):  banben105("G9"):  banben105("E2"):  banben105("F2"):  banben105("G2")
    '  版本10-6  A3  B2  C1
    '  banben106("E9"):  banben106("F9"):  banben106("G9"):  banben106("E2"):  banben106("F2"):  banben106("G2")
    ' '版本13
    'banben131("A11"):  banben131("B11"):  banben131("C11"):  banben131("E11"): banben131("F11"):  banben131("G11"): banben131("I11"):  banben131("J11"):  banben131("K11"):    banben131("M11"):  banben131("N11"):  banben131("O11"):    banben131("Q11"):  banben131("R11"):  banben131("S11"):    banben131("U11"):  banben131("V11"):  banben131("W11"): banben131("Y11"):  banben131("Z11"):  banben131("AA11")
    '
    '  '三个数差1
    'banben131("E6"): banben131("F6"): banben131("G6"): banben131("I6"): banben131("J6"): banben131("K6"): banben131("M6"): banben131("N6"): banben131("O6"): banben131("Q6"): banben131("R6"): banben131("S6"): banben131("U6"): banben131("V6"): banben131("W6"): banben131("Y6"): banben131("Z6"): banben131("AA6")
    'banben131("E7"): banben131("F7"): banben131("G7"): banben131("I7"): banben131("J7"): banben131("K7"): banben131("M7"): banben131("N7"): banben131("O7"): banben131("Q7"): banben131("R7"): banben131("S7"): banben131("U7"): banben131("V7"): banben131("W7"): banben131("Y7"): banben131("Z7"): banben131("AA7")
    'banben131("E8"): banben131("F8"): banben131("G8"): banben131("I8"): banben131("J8"): banben131("K8"): banben131("M8"): banben131("N8"): banben131("O8"): banben131("Q8"): banben131("R8"): banben131("S8"): banben131("U8"): banben131("V8"): banben131("W8"): banben131("Y8"): banben131("Z8"): banben131("AA8")
    'banben131("E9"): banben131("F9"): banben131("G9"): banben131("I9"): banben131("J9"): banben131("K9"): banben131("M9"): banben131("N9"): banben131("O9"): banben131("Q9"): banben131("R9"): banben131("S9"): banben131("U9"): banben131("V9"): banben131("W9"): banben131("Y9"): banben131("Z9"): banben131("AA9")
    'banben131("E10"): banben131("F10"): banben131("G10"): banben131("I10"): banben131("J10"): banben131("K10"): banben131("M10"): banben131("N10"): banben131("O10"): banben131("Q10"): banben131("R10"): banben131("S10"): banben131("U10"): banben131("V10"): banben131("W10"): banben131("Y10"): banben131("Z10"): banben131("AA10")
    'banben131("E11"): banben131("F11"): banben131("G11"): banben131("I11"): banben131("J11"): banben131("K11"): banben131("M11"): banben131("N11"): banben131("O11"): banben131("Q11"): banben131("R11"): banben131("S11"): banben131("U11"): banben131("V11"): banben131("W11"): banben131("Y11"): banben131("Z11"): banben131("AA11")
    '版本21  A-B在绝对值 = B-C 在绝对值
    banben2101("E2") : banben2101("F2") : banben2101("G2")
    banben2101("E3") : banben2101("F3") : banben2101("G3")
    banben2101("E4") : banben2101("F4") : banben2101("G4")
    banben2101("E6") : banben2101("F6") : banben2101("G6")
    banben2101("E7") : banben2101("F7") : banben2101("G7")
    banben2101("E8") : banben2101("F8") : banben2101("G8")
    banben2101("E9") : banben2101("F9") : banben2101("G9")
    banben2101("E10") : banben2101("F10") : banben2101("G10")
    banben2101("E11") : banben2101("F11") : banben2101("G11")
    banben2101("E13") : banben2101("F13") : banben2101("G13")
    banben2101("E14") : banben2101("F14") : banben2101("G14")
    banben2101("E15") : banben2101("F15") : banben2101("G15")
    banben2101("E17") : banben2101("F17") : banben2101("G17")
    banben2101("E18") : banben2101("F18") : banben2101("G18")
    banben2101("E19") : banben2101("F19") : banben2101("G19")
    banben2101("E20") : banben2101("F20") : banben2101("G20")
    banben2101("E21") : banben2101("F21") : banben2101("G21")
    banben2101("E22") : banben2101("F22") : banben2101("G22")
    banben2101("E24") : banben2101("F24") : banben2101("G24")
    banben2101("E25") : banben2101("F25") : banben2101("G25")
    banben2101("E26") : banben2101("F26") : banben2101("G26")
    banben2101("E28") : banben2101("F28") : banben2101("G28")
    banben2101("E29") : banben2101("F29") : banben2101("G29")
    banben2101("E30") : banben2101("F30") : banben2101("G30")
    banben2101("E31") : banben2101("F31") : banben2101("G31")
    banben2101("E32") : banben2101("F32") : banben2101("G32")
    banben2101("E33") : banben2101("F33") : banben2101("G33")
    ''去除低赔率
    If Range("A3").Value < 3 Then
        ' Range("E2:E33,I2:I33,M2:M33,Q2:Q33,U2:U33,Y2:Y33").Interior.Pattern = xlPatternNone
    End If
End Sub
Public Function shuzu(i)
Public Function shuzu1(i)
    Dim X1(18) As Integer
    Dim ran As String
    ran = i
    X1(1) = CInt(Range(ran))
    X1(2) = CInt(Range(ran).Offset(0, 4))
    X1(3) = CInt(Range(ran).Offset(0, 8))
    X1(4) = CInt(Range(ran).Offset(0, 12))
    X1(5) = CInt(Range(ran).Offset(0, 16))
    X1(6) = CInt(Range(ran).Offset(0, 20))
    X1(7) = CInt(Range(ran).Offset(11, 0))
    X1(8) = CInt(Range(ran).Offset(11, 4))
    X1(9) = CInt(Range(ran).Offset(11, 8))
    X1(10) = CInt(Range(ran).Offset(11, 12))
    X1(11) = CInt(Range(ran).Offset(11, 16))
    X1(12) = CInt(Range(ran).Offset(11, 20))
    X1(13) = CInt(Range(ran).Offset(22, 0))
    X1(14) = CInt(Range(ran).Offset(22, 4))
    X1(15) = CInt(Range(ran).Offset(22, 8))
    X1(16) = CInt(Range(ran).Offset(22, 12))
    X1(17) = CInt(Range(ran).Offset(22, 16))
    X1(18) = CInt(Range(ran).Offset(22, 20))
    Dim X11(18) As Integer
    X11(1) = Abs(CInt(Range(ran).Offset(9, 0)))
    X11(2) = Abs(CInt(Range(ran).Offset(9, 4)))
    X11(3) = Abs(CInt(Range(ran).Offset(9, 8)))
    X11(4) = Abs(CInt(Range(ran).Offset(9, 12)))
    X11(5) = Abs(CInt(Range(ran).Offset(9, 16)))
    X11(6) = Abs(CInt(Range(ran).Offset(9, 20)))
    X11(7) = Abs(CInt(Range(ran).Offset(20, 0)))
    X11(8) = Abs(CInt(Range(ran).Offset(20, 4)))
    X11(9) = Abs(CInt(Range(ran).Offset(20, 8)))
    X11(10) = Abs(CInt(Range(ran).Offset(20, 12)))
    X11(11) = Abs(CInt(Range(ran).Offset(20, 16)))
    X11(12) = Abs(CInt(Range(ran).Offset(20, 20)))
    X11(13) = Abs(CInt(Range(ran).Offset(31, 0)))
    X11(14) = Abs(CInt(Range(ran).Offset(31, 4)))
    X11(15) = Abs(CInt(Range(ran).Offset(31, 8)))
    X11(16) = Abs(CInt(Range(ran).Offset(31, 12)))
    X11(17) = Abs(CInt(Range(ran).Offset(31, 16)))
    X11(18) = Abs(CInt(Range(ran).Offset(31, 20)))
    '循环语句
    For x = 1 To 18
        For y = 2 To 18
            If X1(x) = -X1(y) And x <> y Then
                Call getRange(x, ran)
                Call getRange(y, ran)
                If X11(x) / X11(y) = 3 Or X11(y) / X11(x) = 3 Then
                    Call getYanse1(x, ran)
                    Call getYanse1(y, ran)
                End If
            End If
        Next
    Next
End Function
Public Sub getYanse1(x, ran)
    RGB1 = RGB(0, 0, 0)
    If x = 1 Then
        ActiveSheet.Range(ran).Offset(10, 0).Interior.Color = RGB1
    End If
    If x = 2 Then
        ActiveSheet.Range(ran).Offset(10, 4).Interior.Color = RGB1
    End If
    If x = 3 Then
        ActiveSheet.Range(ran).Offset(10, 8).Interior.Color = RGB1
    End If
    If x = 4 Then
        ActiveSheet.Range(ran).Offset(10, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(10, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(10, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(21, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(21, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(21, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(21, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(21, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(21, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(32, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(32, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(32, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(32, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(32, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(32, 20).Interior.Color = RGB1
    End If


End Sub
Public Function shuzu2(i)
    Dim X1(18) As Integer

    Dim ran As String
    ran = i
    X1(1) = CInt(Range(ran))
    X1(2) = CInt(Range(ran).Offset(0, 4))
    X1(3) = CInt(Range(ran).Offset(0, 8))
    X1(4) = CInt(Range(ran).Offset(0, 12))
    X1(5) = CInt(Range(ran).Offset(0, 16))
    X1(6) = CInt(Range(ran).Offset(0, 20))
    X1(7) = CInt(Range(ran).Offset(11, 0))
    X1(8) = CInt(Range(ran).Offset(11, 4))
    X1(9) = CInt(Range(ran).Offset(11, 8))
    X1(10) = CInt(Range(ran).Offset(11, 12))
    X1(11) = CInt(Range(ran).Offset(11, 16))
    X1(12) = CInt(Range(ran).Offset(11, 20))
    X1(13) = CInt(Range(ran).Offset(22, 0))
    X1(14) = CInt(Range(ran).Offset(22, 4))
    X1(15) = CInt(Range(ran).Offset(22, 8))
    X1(16) = CInt(Range(ran).Offset(22, 12))
    X1(17) = CInt(Range(ran).Offset(22, 16))
    X1(18) = CInt(Range(ran).Offset(22, 20))

    Dim X11(18) As Integer

    X11(1) = Abs(CInt(Range(ran).Offset(8, 0)))
    X11(2) = Abs(CInt(Range(ran).Offset(8, 4)))
    X11(3) = Abs(CInt(Range(ran).Offset(8, 8)))
    X11(4) = Abs(CInt(Range(ran).Offset(8, 12)))
    X11(5) = Abs(CInt(Range(ran).Offset(8, 16)))
    X11(6) = Abs(CInt(Range(ran).Offset(8, 20)))
    X11(7) = Abs(CInt(Range(ran).Offset(19, 0)))
    X11(8) = Abs(CInt(Range(ran).Offset(19, 4)))
    X11(9) = Abs(CInt(Range(ran).Offset(19, 8)))
    X11(10) = Abs(CInt(Range(ran).Offset(19, 12)))
    X11(11) = Abs(CInt(Range(ran).Offset(19, 16)))
    X11(12) = Abs(CInt(Range(ran).Offset(19, 20)))
    X11(13) = Abs(CInt(Range(ran).Offset(30, 0)))
    X11(14) = Abs(CInt(Range(ran).Offset(30, 4)))
    X11(15) = Abs(CInt(Range(ran).Offset(30, 8)))
    X11(16) = Abs(CInt(Range(ran).Offset(30, 12)))
    X11(17) = Abs(CInt(Range(ran).Offset(30, 16)))
    X11(18) = Abs(CInt(Range(ran).Offset(30, 20)))

    '循环语句
    For x = 1 To 18
        For y = 2 To 18


            If X1(x) = -X1(y) And x <> y Then


                Call getRange(x, ran)
                Call getRange(y, ran)

                If X11(x) / X11(y) = 3 Or X11(y) / X11(x) = 3 Then

                    Call getYanse2(x, ran)
                    Call getYanse2(y, ran)
                End If

            End If

        Next

    Next

End Function
Public Sub getYanse2(x, ran)
    RGB1 = RGB(0, 0, 0)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(9, 0).Interior.Color = RGB1
    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(9, 4).Interior.Color = RGB1
    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(9, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(9, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(9, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(9, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(20, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(20, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(20, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(20, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(20, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(20, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(31, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(31, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(31, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(31, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(31, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(31, 20).Interior.Color = RGB1
    End If

End Sub

Public Function shuzu3(i)
    Dim X1(18) As Integer

    Dim ran As String
    ran = i
    X1(1) = CInt(Range(ran))
    X1(2) = CInt(Range(ran).Offset(0, 4))
    X1(3) = CInt(Range(ran).Offset(0, 8))
    X1(4) = CInt(Range(ran).Offset(0, 12))
    X1(5) = CInt(Range(ran).Offset(0, 16))
    X1(6) = CInt(Range(ran).Offset(0, 20))
    X1(7) = CInt(Range(ran).Offset(11, 0))
    X1(8) = CInt(Range(ran).Offset(11, 4))
    X1(9) = CInt(Range(ran).Offset(11, 8))
    X1(10) = CInt(Range(ran).Offset(11, 12))
    X1(11) = CInt(Range(ran).Offset(11, 16))
    X1(12) = CInt(Range(ran).Offset(11, 20))
    X1(13) = CInt(Range(ran).Offset(22, 0))
    X1(14) = CInt(Range(ran).Offset(22, 4))
    X1(15) = CInt(Range(ran).Offset(22, 8))
    X1(16) = CInt(Range(ran).Offset(22, 12))
    X1(17) = CInt(Range(ran).Offset(22, 16))
    X1(18) = CInt(Range(ran).Offset(22, 20))

    Dim X11(18) As Integer

    X11(1) = Abs(CInt(Range(ran).Offset(7, 0)))
    X11(2) = Abs(CInt(Range(ran).Offset(7, 4)))
    X11(3) = Abs(CInt(Range(ran).Offset(7, 8)))
    X11(4) = Abs(CInt(Range(ran).Offset(7, 12)))
    X11(5) = Abs(CInt(Range(ran).Offset(7, 16)))
    X11(6) = Abs(CInt(Range(ran).Offset(7, 20)))
    X11(7) = Abs(CInt(Range(ran).Offset(18, 0)))
    X11(8) = Abs(CInt(Range(ran).Offset(18, 4)))
    X11(9) = Abs(CInt(Range(ran).Offset(18, 8)))
    X11(10) = Abs(CInt(Range(ran).Offset(18, 12)))
    X11(11) = Abs(CInt(Range(ran).Offset(18, 16)))
    X11(12) = Abs(CInt(Range(ran).Offset(18, 20)))
    X11(13) = Abs(CInt(Range(ran).Offset(29, 0)))
    X11(14) = Abs(CInt(Range(ran).Offset(29, 4)))
    X11(15) = Abs(CInt(Range(ran).Offset(29, 8)))
    X11(16) = Abs(CInt(Range(ran).Offset(29, 12)))
    X11(17) = Abs(CInt(Range(ran).Offset(29, 16)))
    X11(18) = Abs(CInt(Range(ran).Offset(29, 20)))

    '循环语句
    For x = 1 To 18
        For y = 2 To 18


            If X1(x) = -X1(y) And x <> y Then


                Call getRange(x, ran)
                Call getRange(y, ran)

                If X11(x) / X11(y) = 3 Or X11(y) / X11(x) = 3 Then


                    Call getYanse3(x, ran)
                    Call getYanse3(y, ran)
                End If

            End If

        Next

    Next

End Function
Public Sub getYanse3(x, ran)
    RGB1 = RGB(0, 0, 0)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(8, 0).Interior.Color = RGB1
    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(8, 4).Interior.Color = RGB1
    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(8, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(8, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(8, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(8, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(19, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(19, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(19, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(19, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(19, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(19, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(30, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(30, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(30, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(30, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(30, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(30, 20).Interior.Color = RGB1
    End If


End Sub

Public Sub getRange(x, ran)
    RGB1 = RGB(255, 255, 0)
    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB1
    End If


End Sub
'版本2
Public Function shangxia(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran)))
    X0(2) = Abs(CInt(Range(ran).Offset(0, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(0, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(0, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(0, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(0, 20)))
    X0(7) = Abs(CInt(Range(ran).Offset(11, 0)))
    X0(8) = Abs(CInt(Range(ran).Offset(11, 4)))
    X0(9) = Abs(CInt(Range(ran).Offset(11, 8)))
    X0(10) = Abs(CInt(Range(ran).Offset(11, 12)))
    X0(11) = Abs(CInt(Range(ran).Offset(11, 16)))
    X0(12) = Abs(CInt(Range(ran).Offset(11, 20)))
    X0(13) = Abs(CInt(Range(ran).Offset(22, 0)))
    X0(14) = Abs(CInt(Range(ran).Offset(22, 4)))
    X0(15) = Abs(CInt(Range(ran).Offset(22, 8)))
    X0(16) = Abs(CInt(Range(ran).Offset(22, 12)))
    X0(17) = Abs(CInt(Range(ran).Offset(22, 16)))
    X0(18) = Abs(CInt(Range(ran).Offset(22, 20)))



    Dim Xup(18) As Integer
    Xup(1) = Abs(CInt(Range(ran).Offset(-1, 0)))
    Xup(2) = Abs(CInt(Range(ran).Offset(-1, 4)))
    Xup(3) = Abs(CInt(Range(ran).Offset(-1, 8)))
    Xup(4) = Abs(CInt(Range(ran).Offset(-1, 12)))
    Xup(5) = Abs(CInt(Range(ran).Offset(-1, 16)))
    Xup(6) = Abs(CInt(Range(ran).Offset(-1, 20)))
    Xup(7) = Abs(CInt(Range(ran).Offset(10, 0)))
    Xup(8) = Abs(CInt(Range(ran).Offset(10, 4)))
    Xup(9) = Abs(CInt(Range(ran).Offset(10, 8)))
    Xup(10) = Abs(CInt(Range(ran).Offset(10, 12)))
    Xup(11) = Abs(CInt(Range(ran).Offset(10, 16)))
    Xup(12) = Abs(CInt(Range(ran).Offset(10, 20)))
    Xup(13) = Abs(CInt(Range(ran).Offset(21, 0)))
    Xup(14) = Abs(CInt(Range(ran).Offset(21, 4)))
    Xup(15) = Abs(CInt(Range(ran).Offset(21, 8)))
    Xup(16) = Abs(CInt(Range(ran).Offset(21, 12)))
    Xup(17) = Abs(CInt(Range(ran).Offset(21, 16)))
    Xup(18) = Abs(CInt(Range(ran).Offset(21, 20)))



    Dim XDown(18) As Integer
    XDown(1) = Abs(CInt(Range(ran).Offset(1, 0)))
    XDown(2) = Abs(CInt(Range(ran).Offset(1, 4)))
    XDown(3) = Abs(CInt(Range(ran).Offset(1, 8)))
    XDown(4) = Abs(CInt(Range(ran).Offset(1, 12)))
    XDown(5) = Abs(CInt(Range(ran).Offset(1, 16)))
    XDown(6) = Abs(CInt(Range(ran).Offset(1, 20)))
    XDown(7) = Abs(CInt(Range(ran).Offset(12, 0)))
    XDown(8) = Abs(CInt(Range(ran).Offset(12, 4)))
    XDown(9) = Abs(CInt(Range(ran).Offset(12, 8)))
    XDown(10) = Abs(CInt(Range(ran).Offset(12, 12)))
    XDown(11) = Abs(CInt(Range(ran).Offset(12, 16)))
    XDown(12) = Abs(CInt(Range(ran).Offset(12, 20)))
    XDown(13) = Abs(CInt(Range(ran).Offset(23, 0)))
    XDown(14) = Abs(CInt(Range(ran).Offset(23, 4)))
    XDown(15) = Abs(CInt(Range(ran).Offset(23, 8)))
    XDown(16) = Abs(CInt(Range(ran).Offset(23, 12)))
    XDown(17) = Abs(CInt(Range(ran).Offset(23, 16)))
    XDown(18) = Abs(CInt(Range(ran).Offset(23, 20)))



    '找出第七行绝对值相等 且 上下数的绝对值相等

    For x = 1 To 18
        For y = 2 To 18


            If X0(x) = X0(y) And x <> y Then

                If Xup(x) = Xup(y) And XDown(x) = XDown(y) Then

                    Call getShangxia(x, ran)
                    Call getShangxia(y, ran)

                End If

            End If

        Next

    Next


End Function
Public Sub getShangxia(x, ran)
    RGB1 = RGB(0, 255, 255)
    RGB2 = RGB(0, 255, 255)

    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB2
    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB2


    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB2
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB2
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB2
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB2
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB2
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB2
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB2
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB2
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB2
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB2
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB2
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB2
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB2
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB2
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB2
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB2
    End If


End Sub
'版本3
'第七行 互为相反数  且 上下均为互为相反数
Public Function shangxia7(i)

    Range("D1").Value = i

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = CInt(Range(ran))
    X0(2) = CInt(Range(ran).Offset(0, 4))
    X0(3) = CInt(Range(ran).Offset(0, 8))
    X0(4) = CInt(Range(ran).Offset(0, 12))
    X0(5) = CInt(Range(ran).Offset(0, 16))
    X0(6) = CInt(Range(ran).Offset(0, 20))
    X0(7) = CInt(Range(ran).Offset(11, 0))
    X0(8) = CInt(Range(ran).Offset(11, 4))
    X0(9) = CInt(Range(ran).Offset(11, 8))
    X0(10) = CInt(Range(ran).Offset(11, 12))
    X0(11) = CInt(Range(ran).Offset(11, 16))
    X0(12) = CInt(Range(ran).Offset(11, 20))
    X0(13) = CInt(Range(ran).Offset(22, 0))
    X0(14) = CInt(Range(ran).Offset(22, 4))
    X0(15) = CInt(Range(ran).Offset(22, 8))
    X0(16) = CInt(Range(ran).Offset(22, 12))
    X0(17) = CInt(Range(ran).Offset(22, 16))
    X0(18) = CInt(Range(ran).Offset(22, 20))



    Dim Xup(18) As Integer
    Xup(1) = CInt(Range(ran).Offset(-1, 0))
    Xup(2) = CInt(Range(ran).Offset(-1, 4))
    Xup(3) = CInt(Range(ran).Offset(-1, 8))
    Xup(4) = CInt(Range(ran).Offset(-1, 12))
    Xup(5) = CInt(Range(ran).Offset(-1, 16))
    Xup(6) = CInt(Range(ran).Offset(-1, 20))
    Xup(7) = CInt(Range(ran).Offset(10, 0))
    Xup(8) = CInt(Range(ran).Offset(10, 4))
    Xup(9) = CInt(Range(ran).Offset(10, 8))
    Xup(10) = CInt(Range(ran).Offset(10, 12))
    Xup(11) = CInt(Range(ran).Offset(10, 16))
    Xup(12) = CInt(Range(ran).Offset(10, 20))
    Xup(13) = CInt(Range(ran).Offset(21, 0))
    Xup(14) = CInt(Range(ran).Offset(21, 4))
    Xup(15) = CInt(Range(ran).Offset(21, 8))
    Xup(16) = CInt(Range(ran).Offset(21, 12))
    Xup(17) = CInt(Range(ran).Offset(21, 16))
    Xup(18) = CInt(Range(ran).Offset(21, 20))



    Dim XDown(18) As Integer
    XDown(1) = CInt(Range(ran).Offset(1, 0))
    XDown(2) = CInt(Range(ran).Offset(1, 4))
    XDown(3) = CInt(Range(ran).Offset(1, 8))
    XDown(4) = CInt(Range(ran).Offset(1, 12))
    XDown(5) = CInt(Range(ran).Offset(1, 16))
    XDown(6) = CInt(Range(ran).Offset(1, 20))
    XDown(7) = CInt(Range(ran).Offset(12, 0))
    XDown(8) = CInt(Range(ran).Offset(12, 4))
    XDown(9) = CInt(Range(ran).Offset(12, 8))
    XDown(10) = CInt(Range(ran).Offset(12, 12))
    XDown(11) = CInt(Range(ran).Offset(12, 16))
    XDown(12) = CInt(Range(ran).Offset(12, 20))
    XDown(13) = CInt(Range(ran).Offset(23, 0))
    XDown(14) = CInt(Range(ran).Offset(23, 4))
    XDown(15) = CInt(Range(ran).Offset(23, 8))
    XDown(16) = CInt(Range(ran).Offset(23, 12))
    XDown(17) = CInt(Range(ran).Offset(23, 16))
    XDown(18) = CInt(Range(ran).Offset(23, 20))



    '第七行 互为相反数  且 上下均为互为相反数


    For x = 1 To 18
        For y = 2 To 18


            If X0(x) = -X0(y) And x <> y And X0(x) <> 0 And X0(y) <> 0 Then
                If Xup(x) = -Xup(y) And XDown(x) = -XDown(y) Then
                    If X0(x) * Xup(x) > 0 And X0(x) * XDown(x) > 0 Then

                        getShangxia7(x)
                        getShangxia7(y)

                    End If
                End If
            End If

        Next

    Next



    '第七行 互为相反数   且上下颠倒


    For x = 1 To 18
        For y = 2 To 18


            If X0(x) = -X0(y) And x <> y And X0(x) <> 0 And X0(y) <> 0 Then

                If Xup(x) = XDown(y) And XDown(x) = Xup(y) Then

                    Call getShangxia7(x)
                    Call getShangxia7(y)

                End If
            End If

        Next

    Next


    '版本4
    '找出第七行  七行上=负下一个七行上   七行=下一个七行下    七行下=下一个七行


    For x = 1 To 18
        For y = 2 To 18


            If Xup(x) = -Xup(y) And x <> y And Xup(x) <> 0 Then

                If X0(x) = XDown(y) And XDown(x) = X0(y) Then

                    getSanse(x)
                    getSanse(y)

                End If
            End If

        Next

    Next



    '找出第七行 七行上=负下一个七行 且   七行=下一个七行下     且  七行下=下一个七行上
    '找出第七行 七行上=下一个七行   且   七行=负下一个七行下   且  七行下=下一个七行上
    '找出第七行 七行上=下一个七行   且   七行=下一个七行下     且  七行下=负下一个七行上


    For x = 1 To 18
        For y = 2 To 18


            If Xup(x) = -X0(y) And x <> y And Xup(x) <> 0 Then


                If X0(x) = XDown(y) And XDown(x) = Xup(y) Then

                    getSanseLV(x)
                    getSanseLV(y)

                End If
            End If

        Next

    Next

    '找出第七行 七行上=下一个七行   且   七行=负下一个七行下   且  七行下=下一个七行上

    For x = 1 To 18
        For y = 2 To 18


            If Xup(x) = X0(y) And x <> y And Xup(x) <> 0 Then


                If X0(x) = -XDown(y) And XDown(x) = Xup(y) Then

                    getSanseLV(x)
                    getSanseLV(y)

                End If
            End If

        Next

    Next


    '找出第七行 七行上=下一个七行   且   七行=下一个七行下     且  七行下=负下一个七行上



    For x = 1 To 18
        For y = 2 To 18


            If Xup(x) = X0(y) And x <> y And Xup(x) <> 0 Then
                If X0(x) = XDown(y) And XDown(x) = -Xup(y) Then

                    getSanseLV(x)
                    getSanseLV(y)

                End If
            End If

        Next

    Next





End Function
Public Sub getShangxia7(x)
    ran = Range("D1").Value
    RGB1 = RGB(0, 255, 255)
    RGB2 = RGB(0, 255, 255)
    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB2
    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB2


    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB2
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB2
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB2
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB2
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB2
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB2
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB2
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB2
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB2
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB2
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB2
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB2
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB2
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB2
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB2
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB2
    End If

End Sub
Public Sub getSanse(x)
    ran = Range("D1").Value
    RGB1 = RGB(255, 255, 0)
    RGB2 = RGB(255, 0, 255)
    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB2
    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB2


    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB2
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB2
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB2
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB2
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB2
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB2
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB2
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB2
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB2
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB2
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB2
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB2
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB2
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB2
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB2
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB2
    End If


End Sub
Public Sub getSanseLV(x)
    ran = Range("D1").Value
    RGB1 = RGB(0, 255, 0)
    RGB2 = RGB(0, 255, 0)
    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB2


    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB2

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB2
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB2
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB2
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(-1, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB2
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB2
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB2
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB2
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB2
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB2
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(10, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB2
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 0).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB2
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 4).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB2
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 8).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB2
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 12).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB2
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 16).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB2
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB2
        ActiveSheet.Range(ran).Offset(21, 20).Interior.Color = RGB1
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB2
    End If


End Sub
'版本10-1  A1 B1 C1  绝对值相等
Public Function banben101(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran)))
    X0(2) = Abs(CInt(Range(ran).Offset(0, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(0, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(0, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(0, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(0, 20)))



    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(12, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(12, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(12, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(12, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(12, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(12, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(24, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(24, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(24, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(24, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(24, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(24, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get101(x, ran)
                        Call get101(y, ran)
                        Call get101(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get101(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(24, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(24, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(24, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(24, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(24, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(24, 20).Interior.Color = RGB1
    End If


End Sub
'版本10-2  A1  B3  C2  绝对值相等
Public Function banben102(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran)))
    X0(2) = Abs(CInt(Range(ran).Offset(0, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(0, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(0, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(0, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(0, 20)))


    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(13, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(13, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(13, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(13, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(13, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(13, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(23, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(23, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(23, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(23, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(23, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(23, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get102(x, ran)
                        Call get102(y, ran)
                        Call get102(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get102(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(13, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(13, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(13, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(13, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(13, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(13, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB1
    End If


End Sub
'版本10-3  A2 B1 C3
Public Function banben103(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran).Offset(1, 0)))
    X0(2) = Abs(CInt(Range(ran).Offset(1, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(1, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(1, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(1, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(1, 20)))



    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(11, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(11, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(11, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(11, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(11, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(11, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(24, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(24, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(24, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(24, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(24, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(24, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get103(x, ran)
                        Call get103(y, ran)
                        Call get103(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get103(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(24, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(24, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(24, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(24, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(24, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(24, 20).Interior.Color = RGB1
    End If


End Sub

'版本10-4  A2 B3 C1
Public Function banben104(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran).Offset(1, 0)))
    X0(2) = Abs(CInt(Range(ran).Offset(1, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(1, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(1, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(1, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(1, 20)))


    'B3
    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(13, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(13, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(13, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(13, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(13, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(13, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(22, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(22, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(22, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(22, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(22, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(22, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get104(x, ran)
                        Call get104(y, ran)
                        Call get104(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get104(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(1, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(1, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(1, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(1, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(1, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(13, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(13, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(13, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(13, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(13, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(13, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB1
    End If


End Sub
'版本10-5  A3 B1  C2
Public Function banben105(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran).Offset(2, 0)))
    X0(2) = Abs(CInt(Range(ran).Offset(2, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(2, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(2, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(2, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(2, 20)))



    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(11, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(11, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(11, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(11, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(11, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(11, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(23, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(23, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(23, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(23, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(23, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(23, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get105(x, ran)
                        Call get105(y, ran)
                        Call get105(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get105(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(2, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(2, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(2, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(2, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(2, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(2, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(11, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(11, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(11, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(11, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(11, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(23, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(23, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(23, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(23, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(23, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(23, 20).Interior.Color = RGB1
    End If


End Sub
'版本10-6  A3  B2  C1

Public Function banben106(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran).Offset(2, 0)))
    X0(2) = Abs(CInt(Range(ran).Offset(2, 4)))
    X0(3) = Abs(CInt(Range(ran).Offset(2, 8)))
    X0(4) = Abs(CInt(Range(ran).Offset(2, 12)))
    X0(5) = Abs(CInt(Range(ran).Offset(2, 16)))
    X0(6) = Abs(CInt(Range(ran).Offset(2, 20)))



    Dim X1(18) As Integer
    X1(7) = Abs(CInt(Range(ran).Offset(12, 0)))
    X1(8) = Abs(CInt(Range(ran).Offset(12, 4)))
    X1(9) = Abs(CInt(Range(ran).Offset(12, 8)))
    X1(10) = Abs(CInt(Range(ran).Offset(12, 12)))
    X1(11) = Abs(CInt(Range(ran).Offset(12, 16)))
    X1(12) = Abs(CInt(Range(ran).Offset(12, 20)))


    Dim X2(18) As Integer
    X2(13) = Abs(CInt(Range(ran).Offset(22, 0)))
    X2(14) = Abs(CInt(Range(ran).Offset(22, 4)))
    X2(15) = Abs(CInt(Range(ran).Offset(22, 8)))
    X2(16) = Abs(CInt(Range(ran).Offset(22, 12)))
    X2(17) = Abs(CInt(Range(ran).Offset(22, 16)))
    X2(18) = Abs(CInt(Range(ran).Offset(22, 20)))


    '三行数绝对值相等

    For x = 1 To 6
        For y = 7 To 12

            If X0(x) = X1(y) Then

                For z = 13 To 18

                    If X0(x) = X2(z) Then

                        Call get106(x, ran)
                        Call get106(y, ran)
                        Call get106(z, ran)

                    End If

                Next
            End If

        Next

    Next


End Function
Public Sub get106(x, ran)
    RGB1 = RGB(240, 128, 128)

    If x = 1 Then
        ActiveSheet.Range(ran).Offset(2, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(2, 4).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(2, 8).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(2, 12).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(2, 16).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(2, 20).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(12, 0).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(12, 4).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(12, 8).Interior.Color = RGB1
    End If

    If x = 10 Then
        ActiveSheet.Range(ran).Offset(12, 12).Interior.Color = RGB1
    End If

    If x = 11 Then
        ActiveSheet.Range(ran).Offset(12, 16).Interior.Color = RGB1
    End If

    If x = 12 Then
        ActiveSheet.Range(ran).Offset(12, 20).Interior.Color = RGB1
    End If

    If x = 13 Then
        ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB1
    End If

    If x = 14 Then
        ActiveSheet.Range(ran).Offset(22, 4).Interior.Color = RGB1
    End If

    If x = 15 Then
        ActiveSheet.Range(ran).Offset(22, 8).Interior.Color = RGB1
    End If

    If x = 16 Then
        ActiveSheet.Range(ran).Offset(22, 12).Interior.Color = RGB1
    End If

    If x = 17 Then
        ActiveSheet.Range(ran).Offset(22, 16).Interior.Color = RGB1
    End If

    If x = 18 Then
        ActiveSheet.Range(ran).Offset(22, 20).Interior.Color = RGB1
    End If


End Sub
'三个数绝对值 两个数的合等于第三个数  131 是开头的数组比较
Public Function banben131(i)

    Dim X0(18) As Integer
    Dim ran As String
    ran = i
    X0(1) = Abs(CInt(Range(ran).Offset(0, 0)))
    X0(2) = Abs(CInt(Range(ran).Offset(11, 0)))
    X0(3) = Abs(CInt(Range(ran).Offset(22, 0)))
    'MsgBox X0(1) + X0(2) + X0(3)

    If X0(1) + X0(2) = X0(3) Then

        Call get131(ran)

    End If

    If X0(1) = X0(2) + X0(3) Then

        Call get131(ran)

    End If

    If X0(2) = X0(1) + X0(3) Then

        Call get131(ran)

    End If


    ' 三个数相差1

    ' 1  2  3
    If X0(1) + 1 = X0(2) Then
        If X0(1) + 2 = X0(3) Then

            Call get132(ran)
        End If
    End If

    '   2  1  3
    If X0(1) - 1 = X0(2) Then
        If X0(1) + 1 = X0(3) Then

            Call get132(ran)
        End If
    End If

    '   1  3  2
    If X0(1) + 2 = X0(2) Then
        If X0(1) + 1 = X0(3) Then

            Call get132(ran)
        End If
    End If


    '   2  3  1
    If X0(1) + 1 = X0(2) Then
        If X0(1) - 1 = X0(3) Then

            Call get132(ran)
        End If
    End If

    '3  1  2

    If X0(1) - 2 = X0(2) Then
        If X0(1) - 1 = X0(3) Then

            Call get132(ran)
        End If
    End If

    '3  2  1

    If X0(1) - 1 = X0(2) Then
        If X0(1) - 2 = X0(3) Then

            Call get132(ran)
        End If
    End If

End Function
Public Sub get131(ran)
    RGB1 = RGB(240, 128, 128)

    ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB1



End Sub
Public Sub get132(ran)
    RGB1 = RGB(132, 112, 255)

    ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(11, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(22, 0).Interior.Color = RGB1



End Sub
'版本20  找到93
Public Function banben2001(i)
    Dim X1(10) As Integer
    Dim ran As String
    ran = i
    X1(1) = Abs(CInt(Range(ran).Offset(0, 0)))
    X1(2) = Abs(CInt(Range(ran).Offset(0, 1)))
    X1(3) = Abs(CInt(Range(ran).Offset(0, 2)))
    X1(4) = Abs(CInt(Range(ran).Offset(1, 0)))
    X1(5) = Abs(CInt(Range(ran).Offset(2, 0)))
    X1(6) = Abs(CInt(Range(ran).Offset(1, 1)))
    X1(7) = Abs(CInt(Range(ran).Offset(2, 1)))
    X1(8) = Abs(CInt(Range(ran).Offset(1, 2)))
    X1(9) = Abs(CInt(Range(ran).Offset(2, 2)))
    Dim g As Long
    g = 93

    '循环语句
    For x = 1 To 9
        If X1(x) = g Then
            Call get2001(x, ran)

        End If

    Next
End Function
Public Sub get2001(x, ran)
    RGB1 = RGB(0, 255, 0)
    If x = 1 Then
        ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1

    End If

    If x = 2 Then
        ActiveSheet.Range(ran).Offset(0, 1).Interior.Color = RGB1

    End If

    If x = 3 Then
        ActiveSheet.Range(ran).Offset(0, 2).Interior.Color = RGB1
    End If

    If x = 4 Then
        ActiveSheet.Range(ran).Offset(1, 0).Interior.Color = RGB1
    End If

    If x = 5 Then
        ActiveSheet.Range(ran).Offset(2, 0).Interior.Color = RGB1
    End If

    If x = 6 Then
        ActiveSheet.Range(ran).Offset(1, 1).Interior.Color = RGB1
    End If

    If x = 7 Then
        ActiveSheet.Range(ran).Offset(2, 1).Interior.Color = RGB1
    End If

    If x = 8 Then
        ActiveSheet.Range(ran).Offset(1, 2).Interior.Color = RGB1
    End If

    If x = 9 Then
        ActiveSheet.Range(ran).Offset(2, 2).Interior.Color = RGB1
    End If
End Sub
'版本20  找到31
Public Function banben2002(i)
    Dim X1(10) As Integer
    Dim ran As String
    ran = i
    X1(1) = Abs(CInt(Range(ran).Offset(0, 0)))
    X1(2) = Abs(CInt(Range(ran).Offset(0, 1)))
    X1(3) = Abs(CInt(Range(ran).Offset(0, 2)))
    X1(4) = Abs(CInt(Range(ran).Offset(1, 0)))
    X1(5) = Abs((CInt(Range(ran).Offset(2, 0))))
    X1(6) = Abs(CInt(Range(ran).Offset(1, 1)))
    X1(7) = Abs(CInt(Range(ran).Offset(2, 1)))
    X1(8) = Abs(CInt(Range(ran).Offset(1, 2)))
    X1(9) = Abs(CInt(Range(ran).Offset(2, 2)))
    Dim g As Long
    g = 31

    '循环语句
    For x = 1 To 9
        If X1(x) = g Then
            Call get2001(x, ran)

        End If

    Next
End Function
'版本20  找到0    -1<红0<0
Public Function banben2003(i)
    Dim X1(10) As Double
    Dim ran As String
    ran = i
    X1(1) = CDbl(Range(ran).Offset(0, 0))
    X1(2) = CDbl(Range(ran).Offset(0, 1))
    X1(3) = CDbl(Range(ran).Offset(0, 2))
    X1(4) = CDbl(Range(ran).Offset(1, 0))
    X1(5) = CDbl(Range(ran).Offset(2, 0))
    X1(6) = CDbl(Range(ran).Offset(1, 1))
    X1(7) = CDbl(Range(ran).Offset(2, 1))
    X1(8) = CDbl(Range(ran).Offset(1, 2))
    X1(9) = CDbl(Range(ran).Offset(2, 2))
    '循环语句
    For x = 1 To 9
        If X1(x) < 0 And X1(x) > -0.5 Then
            Call get2001(x, ran)

        End If

    Next
End Function
'版本21   A-B在绝对值=B-C 在绝对值
Public Function banben2101(i)
    Dim X1(10) As Integer
    Dim ran As String
    ran = i
    X1(1) = CInt(Range(ran).Offset(0, 0))
    X1(2) = CInt(Range(ran).Offset(0, 4))
    X1(3) = CInt(Range(ran).Offset(0, 8))
    X1(4) = CInt(Range(ran).Offset(0, 12))
    X1(5) = CInt(Range(ran).Offset(0, 16))
    X1(6) = CInt(Range(ran).Offset(0, 20))

    '    123循环
    '1-2 = 2-3 绝对值
    If Abs(X1(1) - X1(2)) = Abs(X1(2) - X1(3)) And X1(1) <> X1(2) And X1(2) <> X1(3) And X1(1) <> X1(3) Then
        Call get2101(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(1) - X1(2)) = Abs(X1(1) - X1(3)) And X1(1) <> X1(2) And X1(2) <> X1(3) And X1(1) <> X1(3) Then
        Call get2101(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(1) - X1(3)) = Abs(X1(2) - X1(3)) And X1(1) <> X1(2) And X1(2) <> X1(3) And X1(1) <> X1(3) Then
        Call get2101(x, ran)
    End If

    '  234循环
    '1-2 = 2-3 绝对值
    If Abs(X1(2) - X1(3)) = Abs(X1(3) - X1(4)) And X1(2) <> X1(3) And X1(3) <> X1(4) And X1(2) <> X1(4) Then
        Call get2102(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(2) - X1(3)) = Abs(X1(2) - X1(4)) And X1(2) <> X1(3) And X1(3) <> X1(4) And X1(2) <> X1(4) Then
        Call get2102(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(2) - X1(4)) = Abs(X1(3) - X1(4)) And X1(2) <> X1(3) And X1(3) <> X1(4) And X1(2) <> X1(4) Then
        Call get2102(x, ran)
    End If
    '    345循环
    '1-2 = 2-3 绝对值
    If Abs(X1(3) - X1(4)) = Abs(X1(4) - X1(5)) And X1(3) <> X1(4) And X1(4) <> X1(5) And X1(3) <> X1(5) Then
        Call get2103(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(3) - X1(4)) = Abs(X1(3) - X1(5)) And X1(3) <> X1(4) And X1(4) <> X1(5) And X1(3) <> X1(5) Then
        Call get2103(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(3) - X1(5)) = Abs(X1(4) - X1(5)) And X1(3) <> X1(4) And X1(4) <> X1(5) And X1(3) <> X1(5) Then
        Call get2103(x, ran)
    End If
    '456     循环
    '1-2 = 2-3 绝对值
    If Abs(X1(4) - X1(5)) = Abs(X1(5) - X1(6)) And X1(4) <> X1(5) And X1(5) <> X1(6) And X1(4) <> X1(6) Then
        Call get2104(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(4) - X1(5)) = Abs(X1(4) - X1(6)) And X1(4) <> X1(5) And X1(5) <> X1(6) And X1(4) <> X1(6) Then
        Call get2104(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(4) - X1(6)) = Abs(X1(5) - X1(6)) And X1(4) <> X1(5) And X1(5) <> X1(6) And X1(4) <> X1(6) Then
        Call get2104(x, ran)
    End If
    '    561循环
    '1-2 = 2-3 绝对值
    If Abs(X1(5) - X1(6)) = Abs(X1(6) - X1(1)) And X1(5) <> X1(6) And X1(6) <> X1(1) And X1(5) <> X1(1) Then
        Call get2105(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(5) - X1(6)) = Abs(X1(5) - X1(1)) And X1(5) <> X1(6) And X1(6) <> X1(1) And X1(5) <> X1(1) Then
        Call get2105(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(5) - X1(1)) = Abs(X1(6) - X1(1)) And X1(5) <> X1(6) And X1(6) <> X1(1) And X1(5) <> X1(1) Then
        Call get2105(x, ran)
    End If
    '    612循环
    '1-2 = 2-3 绝对值
    If Abs(X1(6) - X1(1)) = Abs(X1(1) - X1(2)) And X1(6) <> X1(1) And X1(1) <> X1(2) And X1(6) <> X1(2) Then
        Call get2106(x, ran)
    End If
    '1-2 = 1-3 绝对值
    If Abs(X1(6) - X1(1)) = Abs(X1(6) - X1(2)) And X1(6) <> X1(1) And X1(1) <> X1(2) And X1(6) <> X1(2) Then
        '        MsgBox (Abs(X1(6) - X1(2)))
        Call get2106(x, ran)
    End If
    '1-3  = 2-3 绝对值
    If Abs(X1(6) - X1(2)) = Abs(X1(1) - X1(2)) And X1(6) <> X1(1) And X1(1) <> X1(2) And X1(6) <> X1(2) Then
        Call get2106(x, ran)
    End If
End Function
Public Sub get2101(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
End Sub
Public Sub get2102(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
End Sub
Public Sub get2103(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 8).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
End Sub
Public Sub get2104(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 12).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
End Sub
Public Sub get2105(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 16).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1
End Sub
Public Sub get2106(x, ran)

    RGB1 = RGB(0, 205, 205)
    ActiveSheet.Range(ran).Offset(0, 20).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 0).Interior.Color = RGB1
    ActiveSheet.Range(ran).Offset(0, 4).Interior.Color = RGB1
End Sub