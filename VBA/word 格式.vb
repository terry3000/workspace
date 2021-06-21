Sub 修改默认格式()

    With ActiveDocument.Styles.Item("标题 1")
        .Font.NameFarEast = "Times New Roman"
        .NextParagraphStyle = "标题 1"

        '字体
        With .Font
            .NameAscii = "Calibri"
            .NameFarEast = "方正小标宋简体"
            .Size = 22
            .Bold = 0                           '1加粗  0不加粗
        End With

        '段落格式
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter  '居中对齐
            .SpaceBefore = 0                     '段前0行
            .SpaceAfter = 0                      '段后0行
            .LineSpacing = 29                    '段落间距
            .CharacterUnitLeftIndent = 0         '文本之前0字符
            .CharacterUnitRightIndent = 0        '文本之后0字符
            .LeftIndent = 0                      '左缩进
            .RightIndent = 0                     '右缩进
            .IndentCharWidth Count:=0            '首行缩进0字符
        End With
    End With


    With ActiveDocument.Styles.Item("标题 2")
        .Font.NameFarEast = "Times New Roman"
        .NextParagraphStyle = "标题 2"

        '字体
        With .Font
            .NameAscii = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = 0                 '1加粗  0不加粗
        End With

        '段落格式
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify    '段落左对齐
            .SpaceBefore = 0                        '段前0行
            .SpaceAfter = 0                         '段后0行
            .LineSpacing = 29                       '段落间距
            .CharacterUnitLeftIndent = 0            '文本之前0字符
            .CharacterUnitRightIndent = 0           '文本之后0字符
            .LeftIndent = 0                         '左缩进
            .RightIndent = 0                        '右缩进
            .CharacterUnitFirstLineIndent = 2       '首行缩进2字符
        End With
    End With

    With ActiveDocument.Styles.Item("正文")
        .Font.NameFarEast = "Times New Roman"
        .NextParagraphStyle = "正文"
                                                    '字体
        With .Font
            .NameAscii = "Calibri"
            .NameFarEast = "仿宋_GB2312"
            .Size = 16
            .Bold = 0                               '1加粗  0不加粗
        End With

                                                    '段落格式
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify    '段落左对齐
            .SpaceBefore = 0                        '段前0行
            .SpaceAfter = 0                         '段后0行
            .LineSpacing = 29                       '段落间距
            .CharacterUnitLeftIndent = 0            '文本之前0字符
            .CharacterUnitRightIndent = 0           '文本之后0字符
            .LeftIndent = 0                         '左缩进
            .RightIndent = 0                        '右缩进
            .CharacterUnitFirstLineIndent = 2       '首行缩进2字符
        End With
    End With



    With ActiveDocument.Styles.Item("标题 3")
        .Font.NameFarEast = "Times New Roman"
        .NextParagraphStyle = "标题 3"

                                                    '字体
        With .Font
            .NameAscii = "Calibri"
            .NameFarEast = "楷体"
            .Size = 16
            .Bold = 0                               '1加粗  0不加粗
        End With

                                                    '段落格式
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify    '段落左对齐
            .SpaceBefore = 0                        '段前0行
            .SpaceAfter = 0                         '段后0行
            .LineSpacing = 29                       '段落间距
            .CharacterUnitLeftIndent = 0            '文本之前0字符
            .CharacterUnitRightIndent = 0           '文本之后0字符
            .LeftIndent = 0                         '左缩进
            .RightIndent = 0                        '右缩进
            .CharacterUnitFirstLineIndent = 2       '首行缩进2字符
        End With
    End With

End Sub
