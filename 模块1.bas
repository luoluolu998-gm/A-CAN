Attribute VB_Name = "模块1"



Option Explicit

Public Sub OpenPhoneForm()
    frmPhone.Show vbModeless
End Sub

Public Sub SortSheet2Columns()

    Dim ws As Worksheet
    Dim col As Long
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim arr As Variant
    Dim temp As Variant
    Dim key1 As Double
    Dim key2 As Double
    Dim swapNeeded As Boolean

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("Sheet2")

    Application.ScreenUpdating = False

    For col = 1 To 20

        lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

        If lastRow >= 3 Then

            arr = ws.Range(ws.Cells(2, col), _
                           ws.Cells(lastRow, col)).value

            For i = 1 To UBound(arr, 1) - 1

                For j = 1 To UBound(arr, 1) - i

                    swapNeeded = False

                    If Trim(CStr(arr(j, 1))) = "" Then

                        If Trim(CStr(arr(j + 1, 1))) <> "" Then
                            swapNeeded = True
                        End If

                    ElseIf Trim(CStr(arr(j + 1, 1))) <> "" Then

                        Select Case col

                            Case 17, 20
                                key1 = ExtractNumber(arr(j, 1))
                                key2 = ExtractNumber(arr(j + 1, 1))

                                If key1 > key2 Then
                                    swapNeeded = True
                                End If

                            Case Else
                                If StrComp(CStr(arr(j, 1)), _
                                           CStr(arr(j + 1, 1)), _
                                           vbTextCompare) > 0 Then
                                    swapNeeded = True
                                End If

                        End Select

                    End If

                    If swapNeeded Then
                        temp = arr(j, 1)
                        arr(j, 1) = arr(j + 1, 1)
                        arr(j + 1, 1) = temp
                    End If

                Next j

            Next i

            ws.Range(ws.Cells(2, col), _
                     ws.Cells(lastRow, col)).value = arr

        End If

    Next col

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "排序时发生错误：" & Err.Description, vbExclamation
    Resume CleanExit

End Sub

Private Function ExtractNumber(ByVal value As Variant) As Double

    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim num As String
    Dim started As Boolean

    s = CStr(value)

    For i = 1 To Len(s)

        ch = Mid(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            num = num & ch
            started = True

        ElseIf started Then
            Exit For
        End If

    Next i

    If num <> "" Then
        ExtractNumber = Val(num)
    Else
        ExtractNumber = 0
    End If

End Function


