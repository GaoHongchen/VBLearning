VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'删除当前sheet中的所有图片
Sub delPics()
Dim pic As Shape
For Each pic In ActiveSheet.Shapes
pic.Delete
Next
End Sub
