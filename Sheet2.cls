VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim cell As Range
    Dim rng As Range
    Dim txt As String

    Set rng = Intersect(Target, Me.Range("A2:S" & Me.Rows.count))

    If rng Is Nothing Then Exit Sub

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    For Each cell In rng.Cells
        If Not cell.HasFormula Then
            If VarType(cell.value) = vbString Then

                ' G至Q列保持原样，不转换大小写
                If cell.Column >= 7 And cell.Column <= 17 Then
                    GoTo NextCell
                End If

                txt = Trim(cell.value)

                ' S列eSIM特殊处理
                If cell.Column = 19 And LCase(txt) = "esim" Then
                    cell.value = "eSIM"
                Else
                    cell.value = UCase(cell.value)
                End If

            End If
        End If
NextCell:
    Next cell

ExitHandler:
    Application.EnableEvents = True

End Sub
