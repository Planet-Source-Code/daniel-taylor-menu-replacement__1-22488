Attribute VB_Name = "OutModule"
Public Function Out(SrcObj As Object, Optional Color1 As OLE_COLOR = -1, Optional Color2 As OLE_COLOR = -1, Optional Text As String = "", Optional TextColor As OLE_COLOR = 0, Optional TextPos As Integer = 5, Optional OutLineText As Boolean = False, Optional OutLineColor As OLE_COLOR = -1, Optional ReverseColor As Boolean = False)
    Dim YPos As Integer, SWidth As Integer, SHeight As Integer
    Dim ScaleModeHolder As Integer
    ScaleModeHolder = SrcObj.ScaleMode
    PrepareObj SrcObj
    'put to vars, faster
    SWidth = SrcObj.ScaleWidth - 1
    SHeight = SrcObj.ScaleHeight - 1
    If Text <> "" Then
        YPos = SrcObj.TextHeight(Text) / 2
    Else
        YPos = 0
    End If
    'check for colors
    If Color1 = -1 Or Color2 = -1 Then
        CheckForColors SrcObj, Color1, Color2
    End If
    If ReverseColor = True Then
        Dim Holder As OLE_COLOR
        Holder = Color1
        Color1 = Color2
        Color2 = Holder
        Holder = Empty
    End If
    'oustide
    SrcObj.Line (0, YPos)-(SWidth, YPos), Color1
    SrcObj.Line (0, YPos)-(0, SHeight), Color1
    SrcObj.Line (0, SHeight)-(SWidth, SHeight), Color2
    SrcObj.Line (SWidth, YPos)-(SWidth, SHeight), Color2
    If Text <> "" Then
        DrawTextForFrame SrcObj, Text, TextPos, Color1, TextColor, OutLineText, OutLineColor
    End If
    SrcObj.ScaleMode = ScaleModeHolder
    End Function


Public Function OutLayered(SrcObj As Object, Times As Integer, Optional Color1 As OLE_COLOR = -1, Optional Color2 As OLE_COLOR = -1, Optional Text As String = "", Optional TextColor As OLE_COLOR = 0, Optional TextPos As Integer = 5, Optional ReverseColor As Boolean = False, Optional OutLineColor As OLE_COLOR = -1)
    On Error GoTo layererror
    'For this function we get the RGB value
    '     of each involved color and
    'fade it into the background color slowl
    '     y, as we move towards the
    'inside.
    'I've finally been able to make this work...
    'looks pretty good now, so you can use it now.
    Dim SWidth As Integer, SHeight As Integer, Count As Integer, YPos As Integer
    Dim Red1 As Integer, Green1 As Integer, Blue1 As Integer
    Dim Red2 As Integer, Green2 As Integer, Blue2 As Integer
    Dim Red3 As Integer, Green3 As Integer, Blue3 As Integer
    Dim Percent As Double, DifR, DifB, DifG, DifR2, DifG2, DifB2
    Dim ScaleModeHolder As Integer
    ScaleModeHolder = SrcObj.ScaleMode
    PrepareObj SrcObj
    'put to vars, faster
    SWidth = SrcObj.ScaleWidth - 1
    SHeight = SrcObj.ScaleHeight - 1
    If Color1 = -1 Or Color2 = -1 Then
        CheckForColors SrcObj, Color1, Color2
    End If
    'If ReverseColor = True Then
    '    Dim Holder As OLE_COLOR
    '    Holder = Color1
    '    Color1 = Color2
    '    Color2 = Holder
    '    Holder = Empty
    'End If
    GetRGB Color1, Red1, Green1, Blue1
    GetRGB Color2, Red2, Green2, Blue2
    GetRGB SrcObj.BackColor, Red3, Green3, Blue3
    'get the diference in color to use later
    DifR = Abs(Red1 - Red3)
    DifG = Abs(Green1 - Green3)
    DifB = Abs(Blue1 - Blue3)
    DifR2 = Abs(Red2 - Red3)
    DifG2 = Abs(Green2 - Green3)
    DifB2 = Abs(Blue2 - Blue3)
    'check if it should be made a frame
    If Text <> "" Then
        YPos = SrcObj.TextHeight(Text) / 2
    Else
        YPos = 0
    End If
    'just draw layer after layer
    For Count = 0 To Times - 1
        Percent = Count / (Times - 1)
        'Debug.Print Percent2
        'get the percent of color mixture betwee
        '     n high/low spots
        'and the backcolor, and use these colors
        '     . increases every
        'time until its the backcolor, supposed
        '     to anyway.....
        If ReverseColor = False Then
            SrcObj.Line (Count, Count + YPos)-(SWidth, Count + YPos), RGB(Red1 - (Percent * DifR), Green1 - (Percent * DifG), Blue1 - (Percent * DifB))
            SrcObj.Line (Count, Count + YPos)-(Count, SHeight), RGB(Red1 - (Percent * DifR), Green1 - (Percent * DifG), Blue1 - (Percent * DifB))
            SrcObj.Line (Count, SHeight)-(SWidth + 1, SHeight), RGB((Percent * DifR2) + Red2, (Percent * DifG2) + Green2, (Percent * DifB2) + Blue2)
            SrcObj.Line (SWidth, Count + YPos)-(SWidth, SHeight + 1), RGB((Percent * DifR2) + Red2, (Percent * DifG2) + Green2, (Percent * DifB2) + Blue2)
        Else
            SrcObj.Line (Count, Count + YPos)-(SWidth, Count + YPos), RGB((Percent * DifR2) + Red2, (Percent * DifG2) + Green2, (Percent * DifB2) + Blue2)
            SrcObj.Line (Count, Count + YPos)-(Count, SHeight), RGB((Percent * DifR2) + Red2, (Percent * DifG2) + Green2, (Percent * DifB2) + Blue2)
            SrcObj.Line (Count, SHeight)-(SWidth + 1, SHeight), RGB(Red1 - (Percent * DifR), Green1 - (Percent * DifG), Blue1 - (Percent * DifB))
            SrcObj.Line (SWidth, Count + YPos)-(SWidth, SHeight + 1), RGB(Red1 - (Percent * DifR), Green1 - (Percent * DifG), Blue1 - (Percent * DifB))
        End If
        SWidth = SWidth - 1
        SHeight = SHeight - 1
    Next Count
    'if its a frame, draw the text
    If Text <> "" Then
        If YPos < Times Then
            If ReverseColor = False Then
                DrawTextForFrame SrcObj, Text, TextPos, Color1, TextColor, True, OutLineColor '
            Else
                DrawTextForFrame SrcObj, Text, TextPos, Color2, TextColor, True, OutLineColor
            End If
        Else
            DrawTextForFrame SrcObj, Text, TextPos, Color1
        End If
    End If
    SrcObj.ScaleMode = ScaleModeHolder
    Exit Function
layererror:
    MsgBox "There was an error. Make sure that Layers <> 0 or 1"
End Function
