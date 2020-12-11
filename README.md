#### **使用VBA代码查找并突出显示Word文档中的重复段落**

若要查找并突出显示Word文档中的重复段落，以下VBA代码可以帮到您，请执行以下操作：

**1**。 按住 **ALT + F11** 键打开 **Microsoft Visual Basic for Applications** 窗口。

**2**。 然后点击 **插页** > **模块**，将下面的代码复制并粘贴到打开的空白模块中：

**VBA代码：查找并突出显示Word文档中的重复段落：**

```vb
Sub highlightdup()
    Dim I, J As Long
    Dim xRngFind, xRng As Range
    Dim xStrFind, xStr As String
    Options.DefaultHighlightColorIndex = wdYellow
    Application.ScreenUpdating = False
    With ActiveDocument
        For I = 1 To .Paragraphs.Count - 1
            Set xRngFind = .Paragraphs(I).Range
            If xRngFind.HighlightColorIndex <> wdYellow Then
                For J = I + 1 To .Paragraphs.Count
                    Set xRng = .Paragraphs(J).Range
                    If xRngFind.Text = xRng.Text Then
                        xRngFind.HighlightColorIndex = wdBrightGreen
                        xRng.HighlightColorIndex = wdYellow
                    End If
                Next
            End If
        Next
    End With
End Sub
```

然后按 **F5** 运行此代码的关键，所有重复的句子一次突出显示，第一个显示的重复段落用绿色突出显示，其他重复段用黄色突出显示。

原网址：http://blog.sina.com.cn/s/blog_6ec242200102z5yd.html



