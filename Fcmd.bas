Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Property Get Commit$()
Commit = WPth & "Commit.Cmd"
End Property
Property Get PushApp$()
PushApp = WPth & "PushApp.Cmd"
End Property