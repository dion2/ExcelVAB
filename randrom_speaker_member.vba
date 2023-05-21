Sub randrom_speaker_member()

 Dim myRange As Range
 Dim check_range As Range
 Dim max_speaker_row As String
 Dim check_speaker_row As String
 Dim ori_speaker_row As String
 
 Set myRange = Worksheets("發表抽獎機").Range("A:A")
 min_value = Application.WorksheetFunction.Min(myRange)
 max_value = Application.WorksheetFunction.Max(myRange)
 '亂數抽籤
 Range("C2").Value = Application.WorksheetFunction.RandBetween(min_value, max_value)
 '找出對應名稱
 Range("D2").Value = Application.WorksheetFunction.VLookup(Range("C2").Value, Range("A:B"), 2, False)

 ori_speaker_row = Application.WorksheetFunction.CountA(Range("A:A"))
 check_speaker_row = Application.WorksheetFunction.CountA(Range("E:E"))
 
 
 If Application.WorksheetFunction.IsNA(Range("G2").Value) Then '檢查該值是否不存在
  '留下已呼叫過的結果
  max_speaker_row = Application.WorksheetFunction.CountA(Range("E:E")) + 1
  Range("E" + max_speaker_row).Value = Range("C2").Value
  Range("F" + max_speaker_row).Value = Range("D2").Value
 ElseIf check_speaker_row = ori_speaker_row Then '檢查是否全數產生過
  MsgBox ("全數已回應")
 Else
  Call randrom_speaker_member '重新遞迴呼叫
 End If
    
End Sub
