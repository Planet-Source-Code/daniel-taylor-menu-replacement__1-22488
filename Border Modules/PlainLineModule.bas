Attribute VB_Name = "PlainLineModule"
Public Function PlainBorder(SrcObj As Object, Optional Color As OLE_COLOR = -1, Optional Width = 1, Optional Text As String = "", Optional TextColor As OLE_COLOR = 0, Optional TextPos As Integer = 5, Optional OutLineText As Boolean = False, Optional OutLineColor As OLE_COLOR = -1)
    'just draw a box around object
    Dim YPos As Integer
    Dim ScaleModeHolder As Integer
    ScaleModeHolder = SrcObj.ScaleMode
    PrepareObj SrcObj
    'check if its supposed to be a frame...
    If Text <> "" Then
        YPos = SrcObj.TextHeight(Text) / 2
    Else
        YPos = 0
    End If
    'colors
    If Color = -1 Then
        CheckForColors SrcObj, , Color
    Else
        TranslateColor Color, 0, Color
    End If
    'if width is 1 then just draw a box, els
    '     e fill the entire thing
    'and delete inside width area
    If Width < 2 Then
        SrcObj.Line (0, YPos)-(SrcObj.ScaleWidth - 1, SrcObj.ScaleHeight - 1), Color, B
    Else
        SrcObj.Line (0, YPos)-(SrcObj.ScaleWidth - 1, SrcObj.ScaleHeight - 1), Color, BF
        SrcObj.Line (Width, YPos + Width)-(SrcObj.ScaleWidth - (1 + Width), SrcObj.ScaleHeight - (1 + Width)), SrcObj.BackColor, BF
    End If
    If Text <> "" Then
        DrawTextForFrame SrcObj, Text, TextPos, Color, TextColor, OutLineText, OutLineColor
    End If
    SrcObj.ScaleMode = ScaleModeHolder
    End Function
