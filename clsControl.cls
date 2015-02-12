VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- clsControl ---

Option Explicit

Public LeftIndent As Currency
Public RightIndent As Currency
Public TopIndent As Currency
Public BottomIndent As Currency
Public Width As Currency
Public Height As Currency
Public ContainerMinWidth As Currency
Public ContainerMinHeight As Currency
Public Style As enLayoutStyle
Public Ctrl As Control
'Public scaleMode



