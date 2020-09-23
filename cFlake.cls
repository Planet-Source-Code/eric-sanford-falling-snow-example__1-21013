VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarX As Long 'local copy
Private mvarY As Long 'local copy
Private mvarOnGround As Boolean 'local copy
Private mvarNextX As Long 'local copy
Private mvarNextY As Long 'local copy


Public Property Let OnGround(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OnGround = 5
    mvarOnGround = vData
End Property


Public Property Get OnGround() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OnGround
    OnGround = mvarOnGround
End Property



Public Property Let y(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get y() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get x() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property



Public Property Let NextY(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarNextY = vData
End Property


Public Property Get NextY() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    NextY = mvarNextY
End Property



Public Property Let NextX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarNextX = vData
End Property


Public Property Get NextX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    NextX = mvarNextX
End Property




