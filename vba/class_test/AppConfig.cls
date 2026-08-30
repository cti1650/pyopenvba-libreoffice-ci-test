VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB_PredeclaredId = True gives this class a global default instance, so
' AppConfig.AppName() works without New. This is the attribute that
' pyOpenVBA preserves verbatim through normalize_class_source().
Option Explicit

Private mName As String

Public Property Get AppName() As String
    If mName = "" Then mName = "PredeclaredDefault"
    AppName = mName
End Property

Public Property Let AppName(ByVal newName As String)
    mName = newName
End Property
