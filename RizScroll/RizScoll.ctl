VERSION 5.00
Begin VB.UserControl RizScroll 
   BackColor       =   &H8000000A&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   KeyPreview      =   -1  'True
   PropertyPages   =   "RizScoll.ctx":0000
   ScaleHeight     =   720
   ScaleWidth      =   735
   ToolboxBitmap   =   "RizScoll.ctx":0019
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   480
   End
   Begin VB.Image Init 
      Height          =   240
      Left            =   240
      Picture         =   "RizScoll.ctx":032B
      Top             =   240
      Width           =   240
   End
   Begin VB.Image MoveRight 
      Height          =   240
      Left            =   480
      Picture         =   "RizScoll.ctx":055D
      Top             =   240
      Width           =   240
   End
   Begin VB.Image MoveLeft 
      Height          =   240
      Left            =   0
      Picture         =   "RizScoll.ctx":0776
      Top             =   240
      Width           =   240
   End
   Begin VB.Image MoveDown 
      Height          =   240
      Left            =   240
      Picture         =   "RizScoll.ctx":0999
      Top             =   480
      Width           =   240
   End
   Begin VB.Image MoveUp 
      Height          =   240
      Left            =   240
      Picture         =   "RizScoll.ctx":0BC2
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "RizScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim Direc As Integer '1 ^ '2 v '3 > '4 <
Private Type TopLeft
    MyTop As Integer
    MyLeft As Integer
End Type

Dim MyTL() As TopLeft

Dim Frm As Object
Dim SS As Integer
'**************************************'
Private Sub MoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Direc = 0: Timer1.Enabled = False
End Sub
Private Sub MoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveUp_Click
        Direc = 1: Timer1.Enabled = True
    End If
End Sub
Private Sub MoveUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveUp_Click
End Sub
Private Sub MoveUp_Click()
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            If Frm.Controls(i).Container Is Frm Then
                Frm.Controls(i).Top = Frm.Controls(i).Top + SS
            End If
        End If
    Next i
End Sub
'**************************************'
Private Sub MoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Direc = 0: Timer1.Enabled = False
End Sub
Private Sub MoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveDown_Click
        Direc = 2: Timer1.Enabled = True
    End If
End Sub
Private Sub MoveDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveDown_Click
End Sub
Private Sub MoveDown_Click()
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            If Frm.Controls(i).Container Is Frm Then
                Frm.Controls(i).Top = Frm.Controls(i).Top - SS
            End If
        End If
    Next i
End Sub
'**************************************'
Private Sub MoveLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Direc = 0: Timer1.Enabled = False
End Sub
Private Sub MoveLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveLeft_Click
        Direc = 4: Timer1.Enabled = True
    End If
End Sub
Private Sub MoveLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveLeft_Click
End Sub
Private Sub MoveLeft_Click()
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            If Frm.Controls(i).Container Is Frm Then
                Frm.Controls(i).Left = Frm.Controls(i).Left - SS
            End If
        End If
    Next i
End Sub
'**************************************'
Private Sub MoveRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Direc = 0: Timer1.Enabled = False
End Sub
Private Sub MoveRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveRight_Click
        Direc = 3: Timer1.Enabled = True
    End If
End Sub
Private Sub MoveRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveRight_Click
End Sub
Private Sub MoveRight_Click()
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            If Frm.Controls(i).Container Is Frm Then
                Frm.Controls(i).Left = Frm.Controls(i).Left + SS
            End If
        End If
    Next i
End Sub
'**************************************'
Private Sub Init_Click()
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            Frm.Controls(i).Left = MyTL(i).MyLeft
            Frm.Controls(i).Top = MyTL(i).MyTop
        End If
    Next i
End Sub
Private Sub Timer1_Timer()
    If Direc = 1 Then Call MoveUp_Click
    If Direc = 2 Then Call MoveDown_Click
    If Direc = 3 Then Call MoveRight_Click
    If Direc = 4 Then Call MoveLeft_Click
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = 720
    UserControl.Height = 720
End Sub
Public Function StartScroll(Form As Object, Optional ScrollSpeed = 10)
    Set Frm = Form
    If IsMissing(ScrollSpeed) Then SS = 10 Else SS = ScrollSpeed
    ReDim MyTL(Frm.Controls.Count - 1) As TopLeft
    On Error Resume Next
    For i = 0 To Frm.Controls.Count - 1
        If Not (TypeOf Frm.Controls(i) Is RizScroll) Then
            MyTL(i).MyLeft = Frm.Controls(i).Left
            MyTL(i).MyTop = Frm.Controls(i).Top
        End If
    Next i
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MoveRight,MoveRight,-1,Enabled
Public Property Get Scrolling_Horizontal() As Boolean
Attribute Scrolling_Horizontal.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Scrolling_Horizontal.VB_ProcData.VB_Invoke_Property = "Properties"
    Scrolling_Horizontal = MoveRight.Enabled
End Property
Public Property Let Scrolling_Horizontal(ByVal New_Scrolling_Horizontal As Boolean)
    MoveRight.Enabled() = New_Scrolling_Horizontal
    MoveLeft.Enabled() = New_Scrolling_Horizontal
    PropertyChanged "Scrolling_Horizontal"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MoveUp,MoveUp,-1,Enabled
Public Property Get Scrolling_Vertical() As Boolean
Attribute Scrolling_Vertical.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Scrolling_Vertical.VB_ProcData.VB_Invoke_Property = "Properties"
    Scrolling_Vertical = MoveUp.Enabled
End Property
Public Property Let Scrolling_Vertical(ByVal New_Scrolling_Vertical As Boolean)
    MoveUp.Enabled() = New_Scrolling_Vertical
    MoveDown.Enabled() = New_Scrolling_Vertical
    PropertyChanged "Scrolling_Vertical"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MoveRight.Enabled = PropBag.ReadProperty("Scrolling_Horizontal", True)
    MoveLeft.Enabled = PropBag.ReadProperty("Scrolling_Horizontal", True)
    MoveUp.Enabled = PropBag.ReadProperty("Scrolling_Vertical", True)
    MoveDown.Enabled = PropBag.ReadProperty("Scrolling_Vertical", True)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Scrolling_Horizontal", MoveRight.Enabled, True)
    Call PropBag.WriteProperty("Scrolling_Vertical", MoveUp.Enabled, True)
    Call PropBag.WriteProperty("Scrolling_Horizontal", MoveLeft.Enabled, True)
    Call PropBag.WriteProperty("Scrolling_Vertical", MoveDown.Enabled, True)
End Sub

