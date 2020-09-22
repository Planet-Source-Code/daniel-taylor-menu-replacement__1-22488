VERSION 5.00
Begin VB.UserControl MenuCtl 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MenuCtl.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "MenuCtl.ctx":08CA
End
Attribute VB_Name = "MenuCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MenuCtl of Menu.ocx
'Created by Daniel Taylor on April 14, 2001
'This is the actual usercontrol, where the user will
'set all the properties later used by MenuFrm.frm
'I have a german version of VB, and most of the code is
'generated, so it has the german comments... just ignore them
'they are for the ActiveX Wizard thing...
'My code is probably very messy & unorganized & unoptimized,
'but the menu is running pretty fast now, almost as fast as
'the windows menus when the mouse if moved over them.
'The code may seem a bit confusing at first, but you need
'to also look at the MenuFrm.frm code to understand it all.
'also the variables are kept in a puclic module (Module1.bas)
'so they can be accessed by the usercontrol and menufrm

Public Enum Style_Type
    Etch_Style
    OutDent_Style
    PlainLine_Style
End Enum
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'Standard-Eigenschaftswerte:
Const m_def_MenuAnimSpeed = 500
Const m_def_OpenAnimated = 0
Const m_def_UseIcons = 0
Const m_def_ItemHotBackColor = &H8000000D
Const m_def_Style = 0
Const m_def_ItemForeColor = &H80000007
Const m_def_ItemHotForeColor = &H8000000E
'Ereignisdeklarationen:
Event ItemClicked(Index As Integer, Text As String)
'api declarations
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 
Private Sub UserControl_Resize()
    'just make the usercontrol a little icon, invisible at runtime
    UserControl.Width = 480
    UserControl.Height = 480
End Sub
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14
Public Function AddItem(Item As String, Optional Enabled As Boolean = True, Optional ItemIcon As StdPicture) As Boolean
    'Add an item to the menu...
    ItemCount = ItemCount + 1
    ReDim Preserve Items(ItemCount)
    Items(ItemCount).Text = Item
    'if it's a seperator disable it, so it is a seperator later
    If LCase(Item) <> "seperator" Then
        Items(ItemCount).Enabled = Enabled
    Else
        Items(ItemCount).Enabled = False
    End If
    'set the icon, if its nothing, its still ok
    Set Items(ItemCount).Pic = ItemIcon
End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14
Public Sub ShowMenu(Optional X As Long = -1, Optional Y As Long = -1)
    'get the mousepos, and set the menufrm.left & .top
    Dim XY As POINTAPI
    Dim LoopMe As Long
    GetCursorPos XY
    Load MenuFrm
    If X = -1 And Y = -1 Then
        MenuFrm.Left = XY.X * Screen.TwipsPerPixelX
        MenuFrm.Top = XY.Y * Screen.TwipsPerPixelY
    Else
        MenuFrm.Left = X * Screen.TwipsPerPixelX
        MenuFrm.Top = Y * Screen.TwipsPerPixelY
    End If
    Set MenuFrm.Font = UserControl.Font
    MenuFrm.Width = 1
    MenuFrm.Height = 1
    'show the form and draw the menu
    MenuFrm.Show
    MenuFrm.DrawMenu , , , True
    MenuClosed = False
    SetCapture MenuFrm.hwnd
    Dim TempText As String, TempIndex As Integer, Raiseevents As Boolean
    'set it into a loop so it checked if the menu is closed,
    'if it is closed, reset the itemdata and raise the
    'itemclick event
    Do
        If MenuClosed = True Then
            'only raise the event it we're on an actual item
            'makes sense because if we are off the form the
            'hotitem is set to -1...
            If HotItem < ItemCount + 2 And HotItem > 0 Then
                'make sure the item is enabled and not a
                'seperator...
                If Items(HotItem - 1).Enabled = True Then
                    TempText = Items(HotItem - 1).Text
                    TempIndex = HotItem
                    Raiseevents = True
                End If
            Else
                Raiseevents = False
            End If
            ItemCount = -1
            ReDim Items(0)
            Exit Do
        End If
        DoEvents
    Loop
    If Raiseevents = True Then
        RaiseEvent ItemClicked(TempIndex, TempText)
    End If
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ItemForeColor() As OLE_COLOR
    ItemForeColor = m_ItemForeColor
End Property

Public Property Let ItemForeColor(ByVal New_ItemForeColor As OLE_COLOR)
    m_ItemForeColor = New_ItemForeColor
    PropertyChanged "ItemForeColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ItemHotForeColor() As OLE_COLOR
    ItemHotForeColor = m_ItemHotForeColor
End Property

Public Property Let ItemHotForeColor(ByVal New_ItemHotForeColor As OLE_COLOR)
    m_ItemHotForeColor = New_ItemHotForeColor
    PropertyChanged "ItemHotForeColor"
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_ItemForeColor = m_def_ItemForeColor
    m_ItemHotForeColor = m_def_ItemHotForeColor
    m_Style = m_def_Style
    m_ItemHotBackColor = m_def_ItemHotBackColor
    m_UseIcons = m_def_UseIcons
    m_BackColor = &H8000000F
    m_OpenAnimated = m_def_OpenAnimated
    m_MenuAnimSpeed = m_def_MenuAnimSpeed
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ItemForeColor = PropBag.ReadProperty("ItemForeColor", m_def_ItemForeColor)
    m_ItemHotForeColor = PropBag.ReadProperty("ItemHotForeColor", m_def_ItemHotForeColor)
    ItemCount = -1
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ItemHotBackColor = PropBag.ReadProperty("ItemHotBackColor", m_def_ItemHotBackColor)
    m_UseIcons = PropBag.ReadProperty("UseIcons", m_def_UseIcons)
    m_BackColor = UserControl.BackColor
    m_OpenAnimated = PropBag.ReadProperty("OpenAnimated", m_def_OpenAnimated)
    m_MenuAnimSpeed = PropBag.ReadProperty("MenuAnimSpeed", m_def_MenuAnimSpeed)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ItemForeColor", m_ItemForeColor, m_def_ItemForeColor)
    Call PropBag.WriteProperty("ItemHotForeColor", m_ItemHotForeColor, m_def_ItemHotForeColor)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ItemHotBackColor", m_ItemHotBackColor, m_def_ItemHotBackColor)
    Call PropBag.WriteProperty("UseIcons", m_UseIcons, m_def_UseIcons)
    Call PropBag.WriteProperty("OpenAnimated", m_OpenAnimated, m_def_OpenAnimated)
    Call PropBag.WriteProperty("MenuAnimSpeed", m_MenuAnimSpeed, m_def_MenuAnimSpeed)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,0
Public Property Get Style() As Variant
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As Variant)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    m_BackColor = New_BackColor
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ItemHotBackColor() As OLE_COLOR
    ItemHotBackColor = m_ItemHotBackColor
End Property

Public Property Let ItemHotBackColor(ByVal New_ItemHotBackColor As OLE_COLOR)
    m_ItemHotBackColor = New_ItemHotBackColor
    PropertyChanged "ItemHotBackColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get UseIcons() As Boolean
Attribute UseIcons.VB_Description = "Use Icons next to text or not?"
    UseIcons = m_UseIcons
End Property

Public Property Let UseIcons(ByVal New_UseIcons As Boolean)
    m_UseIcons = New_UseIcons
    PropertyChanged "UseIcons"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get OpenAnimated() As Boolean
    OpenAnimated = m_OpenAnimated
End Property

Public Property Let OpenAnimated(ByVal New_OpenAnimated As Boolean)
    m_OpenAnimated = New_OpenAnimated
    PropertyChanged "OpenAnimated"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=12,0,0,200
Public Property Get MenuAnimSpeed() As Single
    MenuAnimSpeed = m_MenuAnimSpeed
End Property

Public Property Let MenuAnimSpeed(ByVal New_MenuAnimSpeed As Single)
    m_MenuAnimSpeed = New_MenuAnimSpeed
    PropertyChanged "MenuAnimSpeed"
End Property

