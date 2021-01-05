VERSION 5.00
Begin VB.UserControl MaxMin 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MaxMin.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "MaxMin.ctx":0442
End
Attribute VB_Name = "MaxMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------- API Stuff for border draw -------------

'Win32 API Type declarations
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Win32 API Function declarations
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'Win32 API Constant declarations
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISED = &H5

'------ End API Stuff ---------

'Default Property Values:
Const m_def_MaxPosX = 0
Const m_def_MaxPosY = 0
Const m_def_MaxWidth = 300
Const m_def_MaxHeight = 300
Const m_def_MinWidth = 100
Const m_def_MinHeight = 100
'Property Variables:
Dim m_MaxPosX As Long
Dim m_MaxPosY As Long
Dim m_CenterForm As Boolean
Dim m_MaxWidth As Long
Dim m_MaxHeight As Long
Dim m_MinWidth As Long
Dim m_MinHeight As Long
'Event Declarations:
Event Activate()
Event Deactivate()

Private Sub UserControl_Resize()
    UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
End Sub
Private Sub UserControl_Paint()

    'Draw a 3D raised border on the control using the Win32 API
    Dim rct As RECT

    'First retrieve the control's dimensions into a RECT structure
    GetClientRect UserControl.hwnd, rct

    'Use the DrawEdge Function to draw the 3D border
    DrawEdge UserControl.hdc, rct, BDR_RAISED, BF_RECT

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,300
Public Property Get MaxWidth() As Long
Attribute MaxWidth.VB_Description = "Returns/sets the maximum width of the form."
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Long)
    m_MaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
    MaxX = m_MaxWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,300
Public Property Get MaxHeight() As Long
Attribute MaxHeight.VB_Description = "Returns/sets the maximum height of the form."
    MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Long)
    m_MaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
    MaxY = m_MaxHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get MinWidth() As Long
Attribute MinWidth.VB_Description = "Returns/sets the minimum width of the form."
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
    MinX = m_MinWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_Description = "Returns/sets the minimum height of the form."
    MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    m_MinHeight = New_MinHeight
    PropertyChanged "MinHeight"
    MinY = m_MinHeight
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MaxWidth = m_def_MaxWidth
    m_MaxHeight = m_def_MaxHeight
    m_MinWidth = m_def_MinWidth
    m_MinHeight = m_def_MinHeight
    m_MaxPosX = m_def_MaxPosX
    m_MaxPosY = m_def_MaxPosY
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MaxWidth = PropBag.ReadProperty("MaxWidth", m_def_MaxWidth)
    m_MaxHeight = PropBag.ReadProperty("MaxHeight", m_def_MaxHeight)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
    m_MaxPosX = PropBag.ReadProperty("MaxPosX", m_def_MaxPosX)
    m_MaxPosY = PropBag.ReadProperty("MaxPosY", m_def_MaxPosY)
    
    'Assign values to the public variables used in the subclassing module
    MaxX = m_MaxWidth
    MaxY = m_MaxHeight
    MinX = m_MinWidth
    MinY = m_MinHeight
    MaxPos.x = m_MaxPosX
    MaxPos.y = m_MaxPosY

    If Ambient.UserMode Then
        'reference to this instance of the FormExtender
        Set objFE = Me
        'subclass the parent form
        SubClass UserControl.Extender.Parent.hwnd
    End If
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MaxWidth", m_MaxWidth, m_def_MaxWidth)
    Call PropBag.WriteProperty("MaxHeight", m_MaxHeight, m_def_MaxHeight)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
    Call PropBag.WriteProperty("MaxPosX", m_MaxPosX, m_def_MaxPosX)
    Call PropBag.WriteProperty("MaxPosY", m_MaxPosY, m_def_MaxPosY)
    
    If Ambient.UserMode Then
        'return message control to Windows
        UnSubClass UserControl.Extender.Parent.hwnd
        'Clean up
        Set objFE = Nothing
    End If

End Sub

Friend Sub FormActivated()
    RaiseEvent Activate
End Sub

Friend Sub FormDeActivated()
    RaiseEvent Deactivate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaxPosX() As Long
Attribute MaxPosX.VB_Description = "X position of form when maximised."
    MaxPosX = m_MaxPosX
End Property

Public Property Let MaxPosX(ByVal New_MaxPosX As Long)
    m_MaxPosX = New_MaxPosX
    PropertyChanged "MaxPosX"
    MaxPos.x = m_MaxPosX
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaxPosY() As Long
Attribute MaxPosY.VB_Description = "Y position of form when maximised."
    MaxPosY = m_MaxPosY
End Property

Public Property Let MaxPosY(ByVal New_MaxPosY As Long)
    m_MaxPosY = New_MaxPosY
    PropertyChanged "MaxPosY"
    MaxPos.y = m_MaxPosY
End Property

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    frmAbout.Show 1
End Sub
