Attribute VB_Name = "EtchModule"
Public Function Etch(SrcObj As Object, Optional Color1 As OLE_COLOR = -1, Optional Color2 As OLE_COLOR = -1, Optional Text As String = "", Optional TextColor As OLE_COLOR = 0, Optional TextPos As Integer = 5, Optional OutLineText As Boolean = False, Optional OutLineColor As OLE_COLOR = -1, Optional ReverseColor As Boolean = False)
    Dim YPos As Integer, SWidth As Integer, SHeight As Integer
    Dim ScaleModeHolder As Integer
    ScaleModeHolder = SrcObj.ScaleMode
    PrepareObj SrcObj
    'put to vars, faster
    SWidth = SrcObj.ScaleWidth - 1
    SHeight = SrcObj.ScaleHeight - 1
    'Check if theres text, if so, it's a frame...
    If Text <> "" Then
        YPos = SrcObj.TextHeight(Text) / 2
    Else
        YPos = 0
    End If
    'check what color to use:
    CheckForColors SrcObj, Color1, Color2
    If ReverseColor = True Then
        Dim Holder As OLE_COLOR
        Holder = Color1
        Color1 = Color2
        Color2 = Holder
        Holder = Empty
    End If
    'oustide
    SrcObj.Line (0, YPos)-(SWidth, YPos), Color2
    SrcObj.Line (0, YPos)-(0, SHeight), Color2
    SrcObj.Line (0, SHeight)-(SWidth, SHeight), Color1
    SrcObj.Line (SWidth, YPos)-(SWidth, SHeight), Color1
    'inside
    YPos = YPos + 1
    SWidth = SWidth - 1
    SHeight = SHeight - 1
    SrcObj.Line (1, YPos)-(SWidth, YPos), Color1
    SrcObj.Line (1, YPos)-(1, SHeight), Color1
    SrcObj.Line (1, SHeight)-(SWidth, SHeight), Color2
    SrcObj.Line (SWidth, YPos)-(SWidth, SHeight), Color2
    If Text <> "" Then
        DrawTextForFrame SrcObj, Text, TextPos, Color1, TextColor, OutLineText, OutLineColor
    End If
    SrcObj.ScaleMode = ScaleModeHolder
End Function
