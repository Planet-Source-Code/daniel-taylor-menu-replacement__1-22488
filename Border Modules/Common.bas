Attribute VB_Name = "Common"
'this api call is to get the real color if a system color
'is given in the color argument. i got this off of
'psc under T3D, but my borders can do much more.
Public Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Function GetRGB(Color As OLE_COLOR, Red, Green, Blue)
    'gets Red, Green, and Blue values of a c
    '     olor
    'I think i saw this on www.PlanetSourceC
    '     ode.com
    TranslateColor Color, 0, Color
    Red = Color And &HFF
    Green = (Color And &HFF00&) / 255
    Blue = (Color And &HFF0000) / 65536
End Function

Public Function PrepareObj(SrcObj As Object)
    On Error GoTo PrepareError
    SrcObj.ScaleMode = 3
    SrcObj.AutoRedraw = True
    Exit Function
PrepareError:
    MsgBox "There was an error with the object. Make sure it is a Form/Usercontrol/Picturebox."
End Function

Public Function DrawTextForFrame(SrcObj As Object, Text As String, TextPos As Integer, Color1 As OLE_COLOR, Optional TextColor As OLE_COLOR = -1, Optional OutLineText As Boolean = False, Optional OutLineColor As OLE_COLOR = -1)
    
    Dim ForeCHolder
        'get rid of line where text will be
        SrcObj.Line (TextPos - 1, 0)-(SrcObj.TextWidth(Text) + (TextPos + 1), SrcObj.TextHeight(Text)), SrcObj.BackColor, BF
        'draw the text
        SrcObj.CurrentX = TextPos
        SrcObj.CurrentY = 0

        ForeCHolder = SrcObj.ForeColor
        If TextColor <> -1 Then
            SrcObj.ForeColor = TextColor
        End If
        SrcObj.Print Text
        If OutLineText = True Then
            If OutLineColor = -1 Then
                SrcObj.Line (TextPos - 1, 0)-(SrcObj.TextWidth(Text) + TextPos + 1, SrcObj.TextHeight(Text)), Color1, B
            Else
                SrcObj.Line (TextPos - 1, 0)-(SrcObj.TextWidth(Text) + TextPos + 1, SrcObj.TextHeight(Text)), OutLineColor, B
            End If
        End If
        SrcObj.ForeColor = ForeCHolder
            
End Function

Public Function CheckForColors(SrcObj As Object, Optional Color1 As OLE_COLOR = 0, Optional Color2 As OLE_COLOR = 0)
    If Color1 = -1 Or Color2 = -1 Then
        Dim R As Integer, G As Integer, B As Integer
        GetRGB SrcObj.BackColor, R, G, B
        If Color1 = -1 Then
            If R > 199 Then
                R = 255
            Else
                R = R + 50
            End If
            If G > 199 Then
                G = 255
            Else
                G = G + 50
            End If
            If B > 199 Then
                B = 255
            Else
                B = B + 50
            End If
            Color1 = RGB(R, G, B)
        End If
        If Color2 = -1 Then
            If R < 141 Then
                R = 0
            Else
                R = R - 140
            End If
            If G < 141 Then
                G = 0
            Else
                G = G - 140
            End If
            If B < 141 Then
                B = 0
            Else
                B = B - 140
            End If
            Color2 = RGB(R, G, B)
        End If
    End If
End Function

Public Function GreyOut(SrcObj As Object, Optional Method As Byte = 1, Optional Color As OLE_COLOR = &H808080, Optional Interval As Integer = 2)
    Dim X As Integer, Y As Integer
    SrcObj.ScaleMode = 3
    SrcObj.AutoRedraw = True


    If Method = 1 Then
        'fill regiona with gray dots at interval
        '     s


        For X = 0 To SrcObj.ScaleWidth - 1 Step Interval


            For Y = 0 To SrcObj.ScaleHeight - 1 Step Interval
                SrcObj.PSet (X, Y), Color
            Next Y
        Next X
    Else
        'fill region using grey mask, sometimes


        '     doesn't work...
            Dim DrawModeHolder As Integer
            DrawModeHolder = SrcObj.DrawMode
            SrcObj.DrawMode = 9
            SrcObj.Line (0, 0)-(SrcObj.ScaleWidth, SrcObj.ScaleHeight), Color, BF
            SrcObj.DrawMode = DrawModeHolder
        End If
    End Function


Public Function CText(SrcObj As Object, Text As String, Optional X = "Center", Optional Y = "Center")
    'The easiest way to draw centered text o
    '     n a form/picturebox/ect...
    'You can also supply an X and Y coordina
    '     te to draw at.
    'To use, set the objects font to whateve
    '     r you want and then
    'use CText, it's that easy!
    Dim X1 As Integer, Y1 As Integer
    SrcObj.ScaleMode = 3
    SrcObj.AutoRedraw = True
    X1 = (SrcObj.ScaleWidth / 2) - (SrcObj.TextWidth(Text) / 2)
    Y1 = (SrcObj.ScaleHeight / 2) - (SrcObj.TextHeight(Text) / 2)
    'check if text should be centered or not
    '


    If X = "Center" Then
        SrcObj.CurrentX = X1
    Else
        SrcObj.CurrentX = X
    End If


    If Y = "Center" Then
        SrcObj.CurrentY = Y1
    Else
        SrcObj.CurrentY = Y
    End If
    'finally draw text to control
    SrcObj.Print Text
End Function
