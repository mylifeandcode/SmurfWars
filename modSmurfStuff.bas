Attribute VB_Name = "modSmurfStuff"
Option Explicit
Global SmurfSpeed As Integer
Global SmurfAttack As Integer
Global DamageFactor As Integer
Global BodyCount As Integer
Type tSmurf
     Top As Long
     Left As Long
     Right As Long
     Bottom As Long
     Height As Long
     Width As Long
     Frame As Integer
     Alive As Boolean
End Type

Global Smurf(3) As tSmurf

