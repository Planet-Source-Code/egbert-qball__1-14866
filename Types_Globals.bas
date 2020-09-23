Attribute VB_Name = "Types_Globals"
Option Explicit

Type POINTAPI
X As Long
Y As Long
End Type

Public Type BrickS
X As Long
Y As Long
Height As Long
Width As Long
Visible As Boolean
Pic As Long
End Type

Public Type Bonuss
X As Long
Y As Long
Height As Long
Width As Long
Visible As Boolean
Pic As Long
End Type

Public Type Balls
X As Double
Y As Double
Height As Long
Width As Long
Degree1 As Double
Degree2 As Double
Visible As Boolean
Direction As Long
Pic As Long
End Type

Public Type ShipS
X As Long
Y As Long
Height As Long
Width As Long
StickyBall As Boolean
Cannons As Boolean
HyperSpeed As Boolean
LargeShip As Boolean
End Type

Public Type State
Brick() As BrickS
Ball(1 To 20) As Balls
Ship As ShipS
Bonus(1 To 5) As Bonuss
End Type

Global KeyCodes(0 To 3) As Long
