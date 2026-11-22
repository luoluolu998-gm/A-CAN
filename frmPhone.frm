VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPhone 
   Caption         =   "手机资料录入"
   ClientHeight    =   9576.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11724
   OleObjectBlob   =   "frmPhone.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "frmPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const DB_SHEET As String = "Sheet2"

'==================================================
' 18个普通输入框的名称
'==================================================

Private Function GetControlNames() As Variant

    GetControlNames = Array( _
        "cboBrand", _
        "cboSeries", _
        "cboSuffix", _
        "cboAndroid", _
        "cboScreenSize", _
        "cboScreenType", _
        "cboResolution", _
        "cboCPU", _
        "cboCores", _
        "cboGHz", _
        "cboCortex", _
        "cboStorage", _
        "cboRAM", _
        "cboRearCamera", _
        "cboFrontCamera", _
        "cboBattery", _
        "cboCharging", _
        "cboModelNumber")

End Function


'==================================================
' 18个输入框对应Sheet2的列号
'
' A～Q：17个普通参数
' T：型号数字
'
' R、S列不再使用
'==================================================

Private Function GetColumnNumbers() As Variant

    GetColumnNumbers = Array( _
        1, 2, 3, 4, 5, 6, 7, 8, 9, _
        10, 11, 12, 13, 14, 15, 16, 17, 20)

End Function


Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub lblAndroid_Click()

End Sub

Private Sub lblCPU_Click()

End Sub

Private Sub lblStorage_Click()

End Sub

Private Sub picture2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'==================================================
' 窗体初始化
'==================================================

Private Sub UserForm_Initialize()

    LoadAllLists

    chk5G.value = False
    chkESIM.value = False

    UpdateDeleteButton

End Sub


'==================================================
' 读取A～Q、T列的独立选项
'==================================================

Private Sub LoadAllLists()

    Dim controls As Variant
    Dim columns As Variant
    Dim i As Long

    controls = GetControlNames()
    columns = GetColumnNumbers()

    For i = LBound(controls) To UBound(controls)

        LoadList CStr(controls(i)), CLng(columns(i))

    Next i

End Sub


'==================================================
' 读取指定列的选项
'==================================================

Private Sub LoadList(ByVal controlName As String, _
                     ByVal colNumber As Long)

    Dim ws As Worksheet
    Dim cbo As MSForms.ComboBox
    Dim lastRow As Long
    Dim i As Long
    Dim oldValue As String

    Set ws = ThisWorkbook.Worksheets(DB_SHEET)
    Set cbo = Me.controls(controlName)

    oldValue = CStr(cbo.value)

    cbo.Clear

    lastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row

    For i = 2 To lastRow

        If Trim(CStr(ws.Cells(i, colNumber).value)) <> "" Then

            cbo.AddItem ws.Cells(i, colNumber).value

        End If

    Next i

    ' 保留原来的输入内容
    cbo.value = oldValue

End Sub


'==================================================
' 保存按钮
'
' 只保存18个普通输入框
' 不保存5G和eSIM复选框
'==================================================

Private Sub cmdSavePhone_Click()

    Dim ws As Worksheet
    Dim controls As Variant
    Dim columns As Variant

    Dim i As Long
    Dim valueText As String
    Dim newCount As Long

    On Error GoTo SaveError

    Set ws = ThisWorkbook.Worksheets(DB_SHEET)

    controls = GetControlNames()
    columns = GetColumnNumbers()

    For i = LBound(controls) To UBound(controls)

        valueText = Trim(CStr(Me.controls(controls(i)).value))

        If valueText <> "" Then

            If Not ValueExists(ws, CLng(columns(i)), valueText) Then

                AddValue ws, CLng(columns(i)), valueText

                newCount = newCount + 1

            End If

        End If

    Next i

    LoadAllLists

    UpdateDeleteButton

    MsgBox "保存完成！" & vbCrLf & vbCrLf & _
           "新增选项：" & newCount & " 项" & vbCrLf & _
           "已存在的内容自动跳过。", _
           vbInformation, "保存结果"

    Exit Sub

SaveError:

    MsgBox "保存过程中发生错误：" & vbCrLf & _
           Err.Description, vbExclamation, "保存失败"

End Sub


'==================================================
' 检查对应列是否已经存在
'==================================================

Private Function ValueExists(ByVal ws As Worksheet, _
                             ByVal colNumber As Long, _
                             ByVal searchValue As String) As Boolean

    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row

    For i = 2 To lastRow

        If StrComp(Trim(CStr(ws.Cells(i, colNumber).value)), _
                   searchValue, vbTextCompare) = 0 Then

            ValueExists = True

            Exit Function

        End If

    Next i

End Function


'==================================================
' 将新选项保存到对应列
'==================================================

Private Sub AddValue(ByVal ws As Worksheet, _
                     ByVal colNumber As Long, _
                     ByVal newValue As String)

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row

    If lastRow < 2 Then lastRow = 1

    ws.Cells(lastRow + 1, colNumber).value = newValue

End Sub


'==================================================
' 删除按钮状态判断
'
' 只统计18个普通输入框
' 5G和eSIM不参与判断
'==================================================

Private Sub UpdateDeleteButton()

    Dim controls As Variant
    Dim i As Long
    Dim count As Long

    controls = GetControlNames()

    For i = LBound(controls) To UBound(controls)

        If Trim(CStr(Me.controls(controls(i)).value)) <> "" Then

            count = count + 1

        End If

    Next i

    ' 只有一个普通输入框有内容时才允许删除
    Me.cmdDeleteOption.Enabled = (count = 1)

End Sub


'==================================================
' 删除按钮
'==================================================

Private Sub cmdDeleteOption_Click()

    Dim controls As Variant
    Dim columns As Variant

    Dim i As Long
    Dim selectedIndex As Long

    Dim selectedValue As String
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rowNumber As Long

    Dim answer As VbMsgBoxResult

    controls = GetControlNames()
    columns = GetColumnNumbers()

    selectedIndex = -1

    ' 只检查18个普通输入框
    For i = LBound(controls) To UBound(controls)

        If Trim(CStr(Me.controls(controls(i)).value)) <> "" Then

            If selectedIndex <> -1 Then

                MsgBox "只能删除一个输入框中的内容！", _
                       vbExclamation

                UpdateDeleteButton

                Exit Sub

            End If

            selectedIndex = i

        End If

    Next i

    If selectedIndex = -1 Then

        UpdateDeleteButton

        Exit Sub

    End If

    selectedValue = Trim(CStr( _
        Me.controls(controls(selectedIndex)).value))

    answer = MsgBox( _
        "确定删除这个选项吗？" & vbCrLf & vbCrLf & _
        selectedValue, _
        vbQuestion + vbYesNo, "确认删除")

    If answer <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(DB_SHEET)

    lastRow = ws.Cells(ws.Rows.count, _
                       CLng(columns(selectedIndex))).End(xlUp).Row

    rowNumber = 0

    ' 查找对应列中的选项
    For i = 2 To lastRow

        If StrComp( _
            Trim(CStr(ws.Cells(i, CLng(columns(selectedIndex))).value)), _
            selectedValue, vbTextCompare) = 0 Then

            rowNumber = i

            Exit For

        End If

    Next i

    If rowNumber = 0 Then

        MsgBox "这个内容尚未保存到选项库，无法删除。", _
               vbInformation

        UpdateDeleteButton

        Exit Sub

    End If

    ' 只删除当前列的单元格
    ws.Cells(rowNumber, CLng(columns(selectedIndex))).Delete Shift:=xlUp

    ' 清空当前输入框
    Me.controls(controls(selectedIndex)).value = ""

    ' 重新读取对应下拉列表
    LoadList CStr(controls(selectedIndex)), _
             CLng(columns(selectedIndex))

    Me.controls(controls(selectedIndex)).value = ""

    UpdateDeleteButton

    MsgBox "删除成功！", vbInformation

End Sub


'==================================================
' 18个普通输入框变化时更新删除按钮
'==================================================

Private Sub cboBrand_Change()
    UpdateDeleteButton
End Sub

Private Sub cboSeries_Change()
    UpdateDeleteButton
End Sub

Private Sub cboSuffix_Change()
    UpdateDeleteButton
End Sub

Private Sub cboAndroid_Change()
    UpdateDeleteButton
End Sub

Private Sub cboScreenSize_Change()
    UpdateDeleteButton
End Sub

Private Sub cboScreenType_Change()
    UpdateDeleteButton
End Sub

Private Sub cboResolution_Change()
    UpdateDeleteButton
End Sub

Private Sub cboCPU_Change()
    UpdateDeleteButton
End Sub

Private Sub cboCores_Change()
    UpdateDeleteButton
End Sub

Private Sub cboGHz_Change()
    UpdateDeleteButton
End Sub

Private Sub cboCortex_Change()
    UpdateDeleteButton
End Sub

Private Sub cboStorage_Change()
    UpdateDeleteButton
End Sub

Private Sub cboRAM_Change()
    UpdateDeleteButton
End Sub

Private Sub cboRearCamera_Change()
    UpdateDeleteButton
End Sub

Private Sub cboFrontCamera_Change()
    UpdateDeleteButton
End Sub

Private Sub cboBattery_Change()
    UpdateDeleteButton
End Sub

Private Sub cboCharging_Change()
    UpdateDeleteButton
End Sub


'==================================================
' 型号数字自动转换大写
'==================================================

Private Sub cboModelNumber_Change()

    Static busy As Boolean

    If busy Then Exit Sub

    busy = True

    Dim pos As Long

    pos = cboModelNumber.SelStart

    cboModelNumber.value = UCase(cboModelNumber.value)

    If pos > Len(cboModelNumber.value) Then

        pos = Len(cboModelNumber.value)

    End If

    cboModelNumber.SelStart = pos

    busy = False

    UpdateDeleteButton

End Sub


'==================================================
' GSMArena搜索
'==================================================


Private Sub cmdGSMArena_Click()

    Dim brandName As String
    Dim seriesName As String
    Dim modelNumber As String
    Dim suffixName As String

    Dim lastWord As String
    Dim p As Long

    Dim modelName As String
    Dim searchURL As String

    brandName = Trim(CStr(cboBrand.value))
    seriesName = Trim(CStr(cboSeries.value))
    modelNumber = Trim(CStr(cboModelNumber.value))
    suffixName = Trim(CStr(cboSuffix.value))

    ' 判断系列最后一个单词
    If seriesName <> "" And modelNumber <> "" Then

        p = InStrRev(seriesName, " ")

        If p > 0 Then
            lastWord = Mid(seriesName, p + 1)
        Else
            lastWord = seriesName
        End If

        ' 单个英文字母与型号数字直接连接
        If Len(lastWord) = 1 And lastWord Like "[A-Za-z]" Then

            seriesName = seriesName & modelNumber

        Else

            seriesName = seriesName & " " & modelNumber

        End If

    ElseIf modelNumber <> "" Then

        seriesName = seriesName & " " & modelNumber

    End If

    ' 组合完整型号
    modelName = brandName & " " & seriesName & " " & suffixName

    modelName = Application.Trim(modelName)

    If modelName = "" Then

        MsgBox "请先输入手机品牌和型号。", vbInformation

        Exit Sub

    End If

    searchURL = "https://www.gsmarena.com/results.php3?sQuickSearch=yes&sName=" & _
                Application.EncodeURL(modelName)

    ThisWorkbook.FollowHyperlink Address:=searchURL

End Sub



'==================================================
' 关闭窗体
'==================================================

Private Sub cmdClose_Click()

    Unload Me

End Sub


