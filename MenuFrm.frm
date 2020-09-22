VERSION 5.00
Begin VB.Form MenuFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "MenuFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MenuFrm for Menu.ocx
'Created by Daniel Taylor on April 14, 2001
'This is the actual menu, but we need the usercontrol to
'access and show it through another program.

Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim TPPX As Long, TPPY As Long
    
Private Sub Form_Load()
    HotItem = 0
    OldHotItem = HotItem
    TPPX = Screen.TwipsPerPixelX
    TPPY = Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > MenuFrm.Width Or Y < 0 Or Y > MenuFrm.Height Then
      If HotItem <> -1 Then
        OldHotItem = HotItem
        HotItem = -1
        DrawMenu False, False, False
      End If
    ElseIf CInt(((4 * TPPY) + Y) / TextHeight(Items(0).Text)) <> HotItem And X * TPPX > 0 And X < MenuFrm.Width Then
        OldHotItem = HotItem
        HotItem = CInt(((4 * TPPX) + Y) / TextHeight(Items(0).Text))
        DrawMenu False, False, False
    End If
End Sub

Private Sub Form_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuClosed = True
    ReleaseCapture
    Unload MenuFrm
End Sub

Public Sub DrawMenu(Optional Drawborder As Boolean = True, Optional ResizeMenu As Boolean = True, Optional DrawIcons As Boolean = True, Optional OpenAnim As Boolean = False)
Dim a As Integer, MaxTextWidth As Integer
Dim holdx As Single, holdy As Single
If Drawborder = True Then
    MenuFrm.Cls
End If
MaxTextWidth = 0
MenuFrm.CurrentY = 4 * TPPY
If ItemCount <> -1 Then
  If ResizeMenu = True Then
    For a = 0 To ItemCount
        If TextWidth(Items(a).Text) > MaxTextWidth Then
            MaxTextWidth = TextWidth(Items(a).Text)
        End If
    Next a
    If m_UseIcons = False Then
        MenuFrm.Width = MaxTextWidth + (8 * TPPX)
    Else
        MenuFrm.Width = MaxTextWidth + (26 * TPPX)
    End If
    If m_OpenAnimated = False Then
        MenuFrm.Height = (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
    Else
      If OpenAnim = True Then
        MenuFrm.Height = 10
        For a = 0 To (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
            MenuFrm.Height = MenuFrm.Height + m_MenuAnimSpeed
            If MenuFrm.Height > ((TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)) - m_MenuAnimSpeed Then
                MenuFrm.Height = (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
                DrawMenu
                Exit For
            End If
            DrawMenu , False
            DoEvents
        Next a
      End If
    End If
  End If
  For a = 0 To ItemCount
    If m_UseIcons = False Then
        MenuFrm.CurrentX = 4 * TPPX
    Else
        MenuFrm.CurrentX = (4 + 18) * TPPX
    End If
    On Error Resume Next
    If Items(a).Enabled = True Then
        If a <> HotItem - 1 Then
            holdx = MenuFrm.CurrentX
            holdy = MenuFrm.CurrentY
            If m_UseIcons = True Then
                If DrawIcons = True Then
                    MenuFrm.PaintPicture Items(a).Pic, TPPX * 4, holdy
                    MenuFrm.CurrentX = holdx
                    MenuFrm.CurrentY = holdy
                End If
            End If
            If Drawborder = False Then
                If a = OldHotItem - 1 Then
                    If m_ItemHotBackColor <> m_BackColor Then
                        MenuFrm.Line (holdx - TPPX, holdy)-(MenuFrm.Width - (5 * TPPX), holdy + TextHeight(Items(a).Text)), m_BackColor, BF
                        MenuFrm.CurrentX = holdx
                        MenuFrm.CurrentY = holdy
                    End If
                    MenuFrm.Print Items(a).Text
                Else
                    MenuFrm.CurrentY = MenuFrm.CurrentY + TextHeight(Items(a).Text)
                End If
            Else
                MenuFrm.Print Items(a).Text
            End If
        Else
            holdx = MenuFrm.CurrentX
            holdy = MenuFrm.CurrentY
            If m_UseIcons = True Then
                If DrawIcons = True Then
                    MenuFrm.PaintPicture Items(a).Pic, TPPX * 4, holdy
                    MenuFrm.CurrentX = holdx
                    MenuFrm.CurrentY = holdy
                End If
            End If
            If m_ItemHotBackColor <> m_BackColor Then
                MenuFrm.Line (holdx - TPPX, holdy)-(MenuFrm.Width - (5 * TPPX), holdy + TextHeight(Items(a).Text)), m_ItemHotBackColor, BF
            End If
            MenuFrm.CurrentX = holdx
            MenuFrm.CurrentY = holdy
            MenuFrm.ForeColor = m_ItemHotForeColor
            MenuFrm.Print Items(a).Text
            MenuFrm.ForeColor = m_ItemForeColor
        End If
    Else
      Dim Color1 As OLE_COLOR, Color2 As OLE_COLOR
      Color1 = -1: Color2 = -1
      CheckForColors MenuFrm, Color1, Color2
      If LCase(Items(a).Text) <> "seperator" Then
        holdx = MenuFrm.CurrentX
        holdy = MenuFrm.CurrentY
        If m_UseIcons = True Then
            If DrawIcons = True Then
                MenuFrm.PaintPicture Items(a).Pic, TPPX * 4, holdy
                MenuFrm.CurrentX = holdx
                MenuFrm.CurrentY = holdy
            End If
        End If
        MenuFrm.CurrentX = holdx + (1 * TPPX)
        MenuFrm.CurrentY = holdy + (1 * TPPY)
        MenuFrm.ForeColor = Color1
        MenuFrm.Print Items(a).Text
        MenuFrm.CurrentX = holdx
        MenuFrm.CurrentY = holdy
        MenuFrm.ForeColor = Color2
        MenuFrm.Print Items(a).Text
        MenuFrm.ForeColor = m_ItemForeColor
      Else
        MenuFrm.CurrentX = 4 * TPPX
        holdx = MenuFrm.CurrentX
        holdy = MenuFrm.CurrentY
        MenuFrm.Line (holdx, (holdy + (TextHeight(Items(a).Text) / 2) + 10))-(MenuFrm.Width - (5 * TPPY), (holdy + (TextHeight(Items(a).Text) / 2) + 10)), Color1
        MenuFrm.Line (holdx, holdy + (TextHeight(Items(a).Text) / 2))-(MenuFrm.Width - (5 * TPPY), holdy + (TextHeight(Items(a).Text) / 2)), Color2
        MenuFrm.CurrentX = holdx + TextHeight(Items(a).Text)
        MenuFrm.CurrentY = holdy + TextHeight(Items(a).Text)
      End If
    End If
  Next a
  If Drawborder = True Then
    If m_Style = Etch_Style Then
        Etch MenuFrm
    ElseIf m_Style = OutDent_Style Then
        OutLayered MenuFrm, 3
    Else
        PlainBorder MenuFrm
    End If
  End If
End If
End Sub
