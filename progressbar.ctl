VERSION 5.00
Begin VB.UserControl progressbar 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ScaleHeight     =   300
   ScaleWidth      =   4155
   Begin VB.PictureBox picBar 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim tFillColor As OLE_COLOR 'fill color
Dim lMaxVal As Long         ' max value
Dim lMinVal As Long         'min value
Dim lVal As Long            'value

Const def_max = 100
Const def_min = 0
Const def_fillcolor = vbBlue

' << Max >>
Public Property Let Max(hVal As Long)
lMaxVal = hVal
PropertyChanged "Max"
End Property

Public Property Get Max() As Long
Max = lMaxVal
End Property

' << Min >>
Public Property Let Min(hVal As Long)
lMinVal = hVal
PropertyChanged "Min"
End Property

Public Property Get Min() As Long
Min = lMinVal
End Property

' << Enabled >>
Public Property Let Enabled(hVal As Boolean)
picBar.Enabled = hVal
PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
Enabled = picBar.Enabled
End Property

' << Fill color >>
Public Property Let FillColor(hColor As OLE_COLOR)
tFillColor = hColor
PropertyChanged "FillColor"
End Property

Public Property Get FillColor() As OLE_COLOR
FillColor = tFillColor
End Property

Public Property Let Value(hVal As Long)
lVal = hVal
Call ValChng(hVal)
End Property

Public Property Get Value() As Long
Value = lVal
End Property

Private Sub UserControl_InitProperties()
Max = def_max
Min = def_min
FillColor = def_fillcolor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
tFillColor = PropBag.ReadProperty("FillColor", def_fillcolor)
Enabled = PropBag.ReadProperty("Enabled", True)
Max = PropBag.ReadProperty("Max", def_max)
Min = PropBag.ReadProperty("Min", def_min)
End Sub

Private Sub UserControl_Resize()
picBar.Width = UserControl.Width
picBar.Height = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("FillColor", tFillColor, def_fillcolor)
Call PropBag.WriteProperty("Enabled", Enabled, True)
Call PropBag.WriteProperty("Min", Min, def_min)
Call PropBag.WriteProperty("Max", Max, def_max)
End Sub

Private Sub ValChng(hVal As Long)
Dim caption As String

If hVal > lMaxVal Then
    hVal = lMaxVal
ElseIf hVal < lMinVal Then
    hVal = lMinVal
End If

caption = ((hVal - Min) / (Max - Min)) * 100 & "%"

picBar.Cls
picBar.ScaleWidth = Max - Min
picBar.DrawMode = 10
picBar.CurrentX = (picBar.ScaleWidth / 2 - picBar.TextWidth(caption) / 2)
picBar.CurrentY = (picBar.ScaleHeight - picBar.TextHeight(caption)) / 2

picBar.Print caption
picBar.Line (0, 0)-((hVal - Min), picBar.Width), tFillColor, BF
End Sub
