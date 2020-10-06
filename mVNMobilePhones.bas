Attribute VB_Name = "mVNMobilePhones"
' __   _____   _ ®
' \ \ / / _ | / \
'  \ \ /| _ \/ / \
'   \_/ |___/_/ \_\
'
Option Explicit
#If VBA7 Then
  Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
  Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
  Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
  Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
#End If
#If Win64 Then
  Private gTimerID As LongPtr
#Else
  Private gTimerID As Long
#End If
Private ValArgs(), ValIndex As Integer
Function S_MobileVN(Numbers, _
          Optional Delimiter As String = ",", _
          Optional includeInvalid As Boolean, _
          Optional Expand As Boolean, _
          Optional ZeroFrontNumber As Boolean, _
          Optional Header As Boolean, _
          Optional ReturnOrder As Integer, _
          Optional ReturnNewNumber As Integer, _
          Optional ReturnOldNumber As Integer, _
          Optional ReturnStandardsE164 As Integer, _
          Optional ReturnCompany As Integer, _
          Optional ReturnInvalid As Integer, _
          Optional Title As String) As Variant
      
  On Error Resume Next
  Dim k As Integer, R
  Set R = Application.Caller
  If Title <> vbNullString Then
    S_MobileVN = Title
  Else
    S_MobileVN = Mid(R(1, 1).Formula, 2)
  End If
  k = UBound(ValArgs)
  ReDim Preserve ValArgs(1 To k + 1)
  ValArgs(k + 1) = VBA.Array(R, Numbers, Delimiter, includeInvalid, Expand, ZeroFrontNumber, Header, _
                             ReturnOrder, ReturnNewNumber, ReturnOldNumber, ReturnStandardsE164, ReturnCompany, ReturnInvalid)
  If gTimerID = 0 Then gTimerID = SetTimer(0&, 0&, 1, AddressOf S_ValCallback)
End Function

' __   _____   _ ®
' \ \ / / _ | / \
'  \ \ /| _ \/ / \
'   \_/ |___/_/ \_\
'
Private Sub S_ValCallback()
  On Error Resume Next
  Call KillTimer(0&, gTimerID)
  'On Error GoTo 0
  Dim UA%, S$
  UA = UBound(ValArgs)
  If UA > 0 Then
    ValIndex = ValIndex + 1
    '
    Dim A, B, L1&, L2&, R As Range
    A = ValArgs(ValIndex)
    Set R = A(1)(1, 1)
    L1 = R(A(1).Rows.Count + 2, 1).End(3).Row - R.Row + 1
    If L1 > 0 Then
      B = S_GetMobileVN(R.Resize(L1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12))
      L2 = UBound(B)
    End If
    Call AreaClearContents(A(0)(2, 1), 0, 0, 0, 6)
    If L2 > 0 Then
      A(0)(2, 1).Resize(L2, UBound(B, 2)) = B
    End If
    '
    If ValIndex >= UA Then
      Erase ValArgs: ValIndex = 0
    Else
      gTimerID = SetTimer(0&, 0&, 1, AddressOf S_ValCallback2): Exit Sub
    End If
    Set R = Nothing
  End If
  gTimerID = 0
  On Error GoTo 0
End Sub

Private Sub S_ValCallback2()
  S_ValCallback
End Sub

' __   _____   _ ®
' \ \ / / _ | / \
'  \ \ /| _ \/ / \
'   \_/ |___/_/ \_\
'
Private Sub S_GetMobileVN_test()
  Call S_GetMobileVN("01681234567016812345670168123456701681234567", , , , , 1, 1, 2, 3, 4, 5, 6)
End Sub
Function S_GetMobileVN(ByVal Numbers, _
                   Optional ByVal Delimiter As String = ",", _
                   Optional ByVal includeInvalid As Boolean, _
                   Optional ByVal Expand As Boolean, _
                   Optional ByVal ZeroFrontNumber As Boolean, _
                   Optional ByVal Header As Boolean, _
                   Optional ByVal ReturnOrder As Integer, _
                   Optional ByVal ReturnNewNumber As Integer, _
                   Optional ByVal ReturnOldNumber As Integer, _
                   Optional ByVal ReturnStandardsE164 As Integer, _
                   Optional ByVal ReturnCompany As Integer, _
                   Optional ByVal ReturnInvalid As Integer) As Variant
  Dim nbs, RE, P$, P1$, P2$, P3$, P4$, P5$, S, T, S0$, S1$, S2$, S3$, S4$, S5$, y%, m%
  Dim i As Byte, k&, kk&, c&, L&, F&, Z&, R&, N$, J$, total$(), A(6), ms, m1, m2
  nbs = Numbers: If Not IsArray(nbs) Then nbs = Array(nbs)
  N = vbNullString: m = -Header
  J = IIf(ZeroFrontNumber, "'", N)
  A(1) = ReturnOrder
  A(2) = ReturnNewNumber
  A(3) = ReturnOldNumber
  A(4) = ReturnStandardsE164
  A(5) = ReturnCompany
  A(6) = ReturnInvalid
  If A(1) > y Then y = A(1)
  If A(2) > y Then y = A(2)
  If A(3) > y Then y = A(3)
  If A(4) > y Then y = A(4)
  If A(5) > y Then y = A(5)
  If A(6) > y Then y = A(6)
  Set RE = VBA.CreateObject("VBScript.RegExp")
  With RE
    .IgnoreCase = 1: .Global = 1
    .pattern = "(0|\+84|084|0084)(" & _
      "((?:3[2-9])|(?:86)|(?:9[6-8])" & "|(?:16[2-9]))" & _
      "|((?:7[06-9])|(?:9[03])|(?:89)" & "|(?:12[01268]))" & _
      "|((?:8[1-58])|(?:9[14])" & "|(?:12[34579]))" & _
      "|((?:5[68])|(?:92)" & "|(?:18[68]))" & _
      "|((?:59)|(?:99)" & "|(?:199))" & ")" & _
               "(\d{3})(\d{2})(\d{2})([1-9]*)"
  End With
  If y And Header Then
    k = 1
    ReDim Preserve total(1 To y, 1 To 1):
    If ReturnOrder Then total(ReturnOrder, 1) = "#"
    P1 = "S" & ChrW(7889) & " m" & ChrW(7899) & "i"
    P2 = "S" & ChrW(7889) & " c" & ChrW(361)
    P3 = ChrW(272) & ChrW(7883) & "nh d" & ChrW(7841) & "ng ti" & ChrW(234) & "u chu" & ChrW(7849) & "n"
    P4 = "Nh" & ChrW(224) & " m" & ChrW(7841) & "ng"
    P5 = "Kh" & ChrW(244) & "ng h" & ChrW(7907) & "p l" & ChrW(7879)
    GoSub R
  End If
  For Each S In nbs: GoSub v: Next
  If y Then
    S_GetMobileVN = Application.Transpose(total)
  Else
    S_GetMobileVN = P
  End If
  Set RE = Nothing
Exit Function
v:
  P = S
  With RE
    If y And Not Expand Then GoSub A
    If .Test(S) Then
      P = N: Set ms = .Execute(S)
      For R = 1 To ms.Count
        Set m1 = ms(R - 1): Set m2 = m1.SubMatches
        c = m2.Count
        S0 = m2(0): S1 = m2(1): S2 = m2(c - 4): S3 = m2(c - 3): S4 = m2(c - 2): S5 = m2(c - 1)
        T = Right(S1, 1)
        If y And Expand Then
          F = m1.FirstIndex: Z = m1.Length
          If includeInvalid And (F > L + 2) Then
            GoSub A
            If ReturnInvalid Then total(ReturnInvalid, k) = Mid(S, L + 1, F - L)
          End If
          L = F + Z: GoSub A
        End If

        If ReturnCompany Then
          For i = 2 To 6
            If m2(i) <> N Then
              Select Case i
              Case 2: P = "Viettel"
              Case 3: P = "MobileFone"
              Case 4: P = "VinaPhone"
              Case 5: P = "Vietnamobile"
              Case 6: P = "Beeline/Gmobile"
              Case Else: P = N
              End Select
              Exit For
            End If
          Next
        End If
        Select Case True
        Case S1 Like "16[2-9]": S1 = "3" & T
        Case S1 = "120": S1 = "70"
        Case S1 = "121": S1 = "79"
        Case S1 = "122": S1 = "77"
        Case S1 Like "12[68]": S1 = "7" & T
        Case S1 Like "12[345]": S1 = "8" & T
        Case S1 = "127": S1 = "81"
        Case S1 = "129": S1 = "82"
        Case S1 Like "18[68]": S1 = "5" & T
        Case S1 = "199": S1 = "59"
        End Select
        P1 = P1 & IIf(P1 <> N, Delimiter, J) & "0" & S1 & S2 & S3 & S4
        P2 = P2 & IIf(P2 <> N, Delimiter, J) & IIf(S1 <> CStr(m2(1)), S0 & m2(1) & S2 & S3 & S4, N)
        P3 = P3 & IIf(P3 <> N, Delimiter, N) & "(84)" & S1 & " " & S2 & "-" & S3 & S4
        P4 = P4 & IIf(P4 <> N, Delimiter, N) & P
        P5 = P5 & IIf(P5 <> N, Delimiter, J) & IIf(m2(10) <> vbNullString, S5, N)
        P = P1
        If y And Expand Then GoSub R
      Next
      If y And Not Expand Then GoSub R
    Else
      
    End If
  End With
Return
A:
  kk = kk + 1
  k = kk + m
  ReDim Preserve total(1 To y, 1 To k)
  If ReturnOrder Then total(ReturnOrder, k) = kk
Return
R:
  If ReturnNewNumber Then total(ReturnNewNumber, k) = P1
  If ReturnOldNumber Then total(ReturnOldNumber, k) = P2
  If ReturnStandardsE164 Then total(ReturnStandardsE164, k) = P3
  If ReturnCompany Then total(ReturnCompany, k) = P4
  If ReturnInvalid Then total(ReturnInvalid, k) = P5
  P1 = N: P2 = N: P3 = N: P4 = N: P5 = N
Return
End Function

Sub AreaClearContents(ByVal vRange As Object, Optional ByVal OffsetRow&, Optional ByVal OffsetColumn&, Optional LimitRow&, Optional LimitColumn&)
  Dim R As Object
  Set R = AreaFromTarget(vRange, OffsetRow&, OffsetColumn&, LimitRow, LimitColumn)
  If Not R Is Nothing Then
    R.ClearContents
    Set R = Nothing
  End If
End Sub

Public Function AreaFromTarget(ByVal vRange As Object, _
                    Optional ByVal OffsetRow&, _
                    Optional ByVal OffsetColumn&, _
                    Optional LimitRow&, _
                    Optional LimitColumn&) As Object
  Dim R As Range, T As Range, R1&, C1&, R2&, C2&
  R1 = OffsetRow
  C1 = OffsetColumn
  Set R = vRange(1, 1)
  Set T = R.CurrentRegion
  If T.Cells.Count > 1 Then
    R2 = T.Row + T.Rows.Count - R.Row - R1 + 1
    C2 = T.Column + T.Columns.Count - R.Column - C1 + 1
    If LimitRow > 0 Then
      R2 = IIf(LimitRow < R2, LimitRow, R2)
    End If
    If LimitColumn > 0 Then
      C2 = IIf(LimitColumn < C2, LimitColumn, C2)
    End If
    If R2 > 1 And C2 > 1 Then
      Set AreaFromTarget = R(R1 + 1, C1 + 1).Resize(R2, C2)
    End If
  End If
  Set R = Nothing
  Set T = Nothing
End Function


