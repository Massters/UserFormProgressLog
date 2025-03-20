Attribute VB_Name = "模块1"
Option Explicit

' 模拟文件处理函数 - 外部可调用
Public Sub MockFileProc(filePath As String, sheetName As String)
    ' 模拟工作表处理时间
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' 随机生成错误（10%概率）
    If Rnd() < 0.1 Then
        Err.Raise vbObjectError + 1000, "MockFileProc", "模拟处理错误"
    End If
End Sub

' 启动文件处理器 - 使用无模式显示让用户可以同时操作Excel
Public Sub ShowFileProcessor()
    ' 初始化随机数生成器
    Randomize
    
    ' 无模式显示文件处理窗体，允许用户同时与Excel交互
    FileProcessorForm.Show vbModeless
End Sub

' 工作簿打开时自动显示处理器
Public Sub Auto_Open()
    ShowFileProcessor
End Sub
