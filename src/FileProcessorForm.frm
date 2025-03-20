VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileProcessorForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8090
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   25390
   OleObjectBlob   =   "FileProcessorForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FileProcessorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 模拟文件处理函数 - 实际使用时替换为真实的处理逻辑
Private Sub MockFileProc(filePath As String, sheetName As String)
    ' 模拟工作表处理时间
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' 随机生成错误（10%概率）
    If Rnd() < 0.1 Then
        Err.Raise vbObjectError + 1000, "MockFileProc", "模拟处理错误"
    End If
End Sub

' 表单初始化
Private Sub UserForm_Initialize()
    ' 设置表单位置到屏幕中央
    Me.StartUpPosition = 0 ' 手动
    Me.Caption = "Excel文件处理工具"
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
        
    ' 将Excel窗口移动到屏幕外
    MoveWindowOffScreen
End Sub

' 表单关闭事件
Private Sub UserForm_Terminate()
    ' 恢复Excel窗口到原始位置
'    RestoreWindowPosition
    ThisWorkbook.Close savechanges:=False
End Sub

' 将Excel窗口移动到屏幕外
Private Sub MoveWindowOffScreen()
    With Application
        .WindowState = xlNormal
        .Left = 10000 ' 移到屏幕左侧外
        .Top = 0
    End With
End Sub

' 恢复Excel窗口到原始位置
Private Sub RestoreWindowPosition()
    Application.WindowState = xlMaximized
End Sub

' 选择文件按钮点击事件
Private Sub btnSelect_Click()
    Dim fileDialog As fileDialog
    Dim selectedItem As Variant
    Dim filePaths As String
    
    ' 创建文件对话框
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .AllowMultiSelect = True
        .Title = "请选择Excel文件"
        .Filters.Add "Excel文件", "*.xlsx; *.xls", 1
        
        ' 如果用户选择了文件并点击了"确定"
        If .Show = -1 Then
            ' 清空当前文件路径
            filePaths = ""
            
            ' 遍历所有选中的文件
            For Each selectedItem In .SelectedItems
                filePaths = filePaths & selectedItem & vbCrLf
            Next selectedItem
            
            ' 更新文件路径文本框（移除最后一个换行符）
            If Len(filePaths) > 0 Then
                filePaths = Left(filePaths, Len(filePaths) - Len(vbCrLf))
            End If
            txtFilePath.Text = filePaths
        End If
    End With
    
    Set fileDialog = Nothing

End Sub

' 开始处理按钮点击事件
Private Sub btnStart_Click()
    Dim filePaths() As String
    Dim filePath As String
    Dim fileName As String
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim errorMessage As String
    Dim currentWb As Workbook
    Dim result As String
    Dim originalCalculation As XlCalculation
    
    ' 保存当前工作簿的引用
    Set currentWb = ThisWorkbook
    
    ' 保存当前的计算模式
    originalCalculation = Application.Calculation
    
    ' 清空进度信息
    txtProgress.Text = ""
    
    ' 如果没有选择文件，则退出
    If Trim(txtFilePath.Text) = "" Then
        MsgBox "请先选择文件！", vbExclamation
        Exit Sub
    End If
    
    ' 分割文件路径
    filePaths = Split(txtFilePath.Text, vbCrLf)
    
    ' 禁用按钮，防止重复点击
    btnSelect.Enabled = False
    btnStart.Enabled = False
    
    ' 优化性能设置
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' 确保Excel窗口在屏幕外
    MoveWindowOffScreen
    
    ' 处理每个文件
    For i = 0 To UBound(filePaths)
        filePath = Trim(filePaths(i))
        
        ' 获取文件名（不含路径）
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        
        ' 更新进度信息
        AppendProgressInfo fileName
        
        ' 尝试打开文件
        On Error Resume Next
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        
        ' 检查是否有错误
        If Err.Number <> 0 Then
            ' 记录文件打开错误
            errorMessage = "Error: " & Err.Description
            AppendProgressInfo String(8, ".") & errorMessage
            
            Err.Clear
            On Error GoTo 0
        Else
            ' 文件打开成功，处理每个工作表
            On Error GoTo 0
            
            ' 处理工作簿中的每个工作表
            For Each ws In wb.Worksheets
                ' 尝试处理工作表
                On Error Resume Next
                MockFileProc filePath, ws.Name
                
                ' 检查是否有错误
                If Err.Number <> 0 Then
                    ' 记录处理错误
                    result = "Error: " & Err.Description
                    Err.Clear
                Else
                    result = "Done"
                End If
                On Error GoTo 0
                
                ' 直接输出最终状态（不显示ing状态）
                AppendProgressInfo "    " & ws.Name & String(8, ".") & result
            Next ws
            
            ' 关闭工作簿
            wb.Close savechanges:=False
        End If
        
        Set wb = Nothing
    Next i
    
    ' 恢复原始设置
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = originalCalculation
    
    ' 启用按钮
    btnSelect.Enabled = True
    btnStart.Enabled = True
    
    ' 确保窗体在前面
'    Me.SetFocus
    
    MsgBox "所有文件处理完成！", vbInformation
End Sub

' 追加进度信息
Private Sub AppendProgressInfo(info As String)
    ' 追加信息到进度文本框
    txtProgress.Text = txtProgress.Text & info & vbCrLf
    
    ' 滚动到最后一行
    txtProgress.SelStart = Len(txtProgress.Text)
    
    ' 更新界面
    Me.Repaint
'    DoEvents
End Sub

