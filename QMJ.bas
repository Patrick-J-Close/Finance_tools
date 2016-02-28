Attribute VB_Name = "QMJ"
Option Explicit
'
Private ws_name As String
Private factor As String
Private lastRow As Long
'
Public Sub QMJ()
    '
    Dim ws() As Variant
    Dim i As Integer
    '
    Application.ScreenUpdating = False
    '
    ws = GetSheets
    '
    'for loop begins at second full year
    For i = 5 To UBound(ws)
        '
        ws_name = ws(i)
        lastRow = Sheets(ws_name).Range("A1").End(xlDown).Row
        Call cleanNAs
        '
        factor = "PROF"
        Call CalcProf
        'CalcGrowth is called within CalcProf becuase of data array overlap and thus is
        'not called by the master sub
        '
        factor = "SAFE"
        Call CalcSafety
        '
        factor = "PAYO"
        Call CalcPayout
        '
    Next i
    '
    Application.ScreenUpdating = True
    '
End Sub
'
Private Function CalcProf()
    '
    Dim GPOA_f() As String, f() As String, CFOA_f() As String
    Dim i As Integer
    '
    Call AddSheetandTickers
    '
    'Copy & Paste ROA, ROE, GM
    f = Split("RETURN_COM_EQY,RETURN_ON_ASSET,GROSS_MARGIN", ",")
    For i = 0 To 2
        Call fieldCopy(f(i), ws_name, lastRow, i)
    Next i
    '
    'GPOA = (gross profit) / (total assets)
    Dim REV() As Variant, GM() As Variant, TA() As Variant
    GPOA_f = Split("SALES_REV_TURN,GROSS_MARGIN,BS_TOT_ASSET", ",")
    REV = fieldArray(GPOA_f(0), ws_name, lastRow)
    GM = fieldArray(GPOA_f(1), ws_name, lastRow)
    TA = fieldArray(GPOA_f(2), ws_name, lastRow)
    '
    Sheets(ws_name & "_PROF").Activate
    Range("E1").Value = "GPOA"
    '
    On Error Resume Next
    For i = 1 To lastRow
        '
        Range("A1").Offset(i, 4).Value = (REV(i, 1) * (GM(i, 1) / 100)) / TA(i, 1)
    Next i
    On Error GoTo 0
    '
    'CFOA = (net income + D&A - chg in WC - CapEx) / (total assets)
    'Note: CapEx is already negative and is thus added
    Dim NI() As Variant, DA() As Variant, chgWC() As Variant, CapEx() As Variant
    CFOA_f = Split("NET_INCOME,CF_DEPR_AMORT,CHNG_WORK_CAP,CAPITAL_EXPEND", ",")
    NI = fieldArray(CFOA_f(0), ws_name, lastRow)
    DA = fieldArray(CFOA_f(1), ws_name, lastRow)
    chgWC = fieldArray(CFOA_f(2), ws_name, lastRow)
    CapEx = fieldArray(CFOA_f(3), ws_name, lastRow)
    '
    Sheets(ws_name & "_PROF").Activate
    Range("F1").Value = "CFOA"
    On Error Resume Next
    For i = 1 To lastRow
        '
        Range("A1").Offset(i, 5).Value = _
        (NI(i, 1) + DA(i, 1) - chgWC(i, 1) + CapEx(i, 1)) / TA(i, 1)
    Next i
    On Error GoTo 0
    '
    'ACC = (D&A - chg in WC)/ (total assets)
    Range("G1").Value = "ACC"
    On Error Resume Next
    For i = 1 To lastRow
        '
        Range("A1").Offset(i, 6).Value = _
        (DA(i, 1) - chgWC(i, 1)) / TA(i, 1)
    Next i
    On Error GoTo 0
    '
    Call CalcZscores(ws_name & "_PROF", lastRow)
    Call CalcGrowth(GPOA_f, CFOA_f, REV, GM, NI, DA, chgWC, CapEx)
    '
End Function
'
Private Function CalcGrowth(GPOA_f() As String, CFOA_f() As String, REV() As Variant, GM() As Variant, _
    NI() As Variant, DA() As Variant, chgWC() As Variant, CapEx() As Variant)
    '
    Dim i As Integer
    Dim s() As Variant
    '
    factor = "GROW"
    Call AddSheetandTickers
    s = GetTickers(ws_name)
    '
    'Delta GPOA = (GP - GP(t-5)) / TA(t-5), where GP = REV * GM
    Dim REV_t5 As Variant, GM_t5() As Variant, TA_t5() As Variant
    REV_t5 = fieldarrayHist(s, GPOA_f(0), ws_name, 1)   '###CHANGE to 5####
    GM_t5 = fieldarrayHist(s, GPOA_f(1), ws_name, 1)
    TA_t5 = fieldarrayHist(s, GPOA_f(2), ws_name, 1)
    '
    Sheets(ws_name & "_GROW").Activate
    Range("B1").Value = "DEL_GPOA"
    '
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 1).Value = _
        ((REV(i + 1, 1) * (GM(i + 1, 1) / 100)) - (REV_t5(i) * (GM_t5(i) / 100))) / TA_t5(i)
    Next i
    On Error GoTo 0
    '
    'Delta CFOA = (net income + D&A - chg in WC - CapEx) / (total assets)
    'note: CapEX is already negative and is thus added
    Dim NI_t5() As Variant, DA_t5() As Variant, chgWC_t5() As Variant, CapEx_t5() As Variant
    NI_t5 = fieldarrayHist(s, CFOA_f(0), ws_name, 1)
    DA_t5 = fieldarrayHist(s, CFOA_f(1), ws_name, 1)
    chgWC_t5 = fieldarrayHist(s, CFOA_f(2), ws_name, 1)    '###CHANGE to 5####
    CapEx_t5 = fieldarrayHist(s, CFOA_f(3), ws_name, 1)
    '
    Sheets(ws_name & "_GROW").Activate
    Range("C1").Value = "DEL_CFOA"
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 2).Value = _
        ((NI(i + 1, 1) + DA(i + 1, 1) - chgWC(i + 1, 1) + CapEx(i + 1, 1)) _
        - (NI_t5(i) + DA_t5(i) - chgWC_t5(i) + CapEx_t5(i))) / TA_t5(i)
    Next i
    On Error GoTo 0
    '
    'Delta ROE (5 year) = (NI - NI_t5) / (Book value of Equity  in t-5)
    'substituting ROE for BV Equity: ((NI - NI_t5) * ROE_t5) / NI_t5
    Dim ROE_t5() As Variant
    ROE_t5 = fieldarrayHist(s, "RETURN_COM_EQY", ws_name, 1)   '###CHANGE to 5####
    '
    Sheets(ws_name & "_GROW").Activate
    Range("D1").Value = "DEL_ROE"
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 3).Value = _
        ((NI(i + 1, 1) - NI_t5(i)) * ROE_t5(i)) / NI_t5(i)
    Next i
    On Error GoTo 0
    '
    'Delta ROA (5 year) = (NI - NI_t5) / (total assets in t-5)
    Sheets(ws_name & "_GROW").Activate
    Range("E1").Value = "DEL_ROA"
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 4).Value = _
        (NI(i + 1, 1) - NI_t5(i)) / TA_t5(i)
    Next i
    On Error GoTo 0
    '
    'Delta Gross Margin (5 year) = (GP - GP_t5) / REV_t5
    ' where GP = REV * GM
    Sheets(ws_name & "_GROW").Activate
    Range("E1").Value = "DEL_GM"
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 5).Value = _
        ((REV(i + 1, 1) * (GM(i + 1, 1) / 100)) - (REV_t5(i) * (GM_t5(i) / 100))) / REV_t5(i)
    Next i
    On Error GoTo 0
    '
    'Delta ACC (5 year) = ((DA - chgWC) - (DA_t5 - chgWC_t5)) / TA_t5
    Sheets(ws_name & "_GROW").Activate
    Range("f1").Value = "DEL_ACC"
    On Error Resume Next
    For i = 0 To lastRow
        '
        Range("A2").Offset(i, 6).Value = _
        ((DA(i, 1) - chgWC(i, 1)) - (DA_t5(i) - chgWC_t5(i))) / TA_t5(i)
    Next i
    On Error GoTo 0
    '
    Call CalcZscores(ws_name & "_GROW", lastRow)
    '
End Function

Private Function CalcSafety()
    '
    Dim i As Integer
    Dim f() As String
    '
    Call AddSheetandTickers
    '
    'Copy & Paste Beta, volatility, D/E, Altman's Z
    f = Split("EQY_BETA,VOLATILITY_360D,TOT_DEBT_TO_COM_EQY,ALTMAN_Z_SCORE", ",")
    For i = 0 To 2
        Call fieldCopy(f(i), ws_name, lastRow, i)
    Next i
    '
    Call CalcZscores(ws_name & "_SAFE", lastRow)
    '
End Function
'
Private Function CalcPayout()
'
    Dim s() As Variant
    Dim i As Integer
    Dim EISS_f As String, DISS_f() As String
    '
    Call AddSheetandTickers
    s = GetTickers(ws_name)
    '
    'EISS = -(sharesOut/sharesOut_t-1)
    Dim shOut() As Variant, shOut_t1() As Variant
    EISS_f = "IS_SH_FOR_DILUTED_EPS"
    shOut = fieldArray(EISS_f, ws_name, lastRow)
    shOut_t1 = fieldarrayHist(s, EISS_f, ws_name, 1)
    '
    Sheets(ws_name & "_PAYO").Activate
    Range("B1").Value = "EISS"
    On Error Resume Next
    For i = 0 To UBound(shOut)
        Range("A2").Offset(i, 1).Value = -shOut(i + 1, 1) / shOut_t1(i)
    Next i
    On Error GoTo 0
    '
    'DISS = -((ST & LT Debt + minority int + preferred equity)/(same quantity one year prior)
    'DISS = -(one year % change in total debt)
    Dim TD() As Variant, TD_t1() As Variant, PEMI() As Variant, PEMI_t1() As Variant
    DISS_f = Split("SHORT_AND_LONG_TERM_DEBT,PREFERRED_EQUITY_&_MINORITY_INT", ",")
    TD = fieldArray(DISS_f(0), ws_name, lastRow)
    PEMI = fieldArray(DISS_f(1), ws_name, lastRow)
    TD_t1 = fieldarrayHist(s, DISS_f(0), ws_name, 1)
    PEMI_t1 = fieldarrayHist(s, DISS_f(1), ws_name, 1)
    '
    Sheets(ws_name & "_PAYO").Activate
    Range("C1").Value = "DISS"
    On Error Resume Next
    For i = 0 To UBound(TD)
        Range("A2").Offset(i, 2).Value = -(TD(i + 1, 1) + PEMI(i + 1, 1)) / (TD_t1(i) + PEMI_t1(i))
    Next i
    On Error GoTo 0
    '
    Call CalcZscores(ws_name & "_PAYO", lastRow)
    
End Function
'
Private Function CalcZscores(ByRef ws As String, ByVal lastRow As Long)
'
    Dim ColNum As Integer, i As Integer, j As Integer
    '
    Sheets(ws).Activate
    ColNum = Range("A1", Range("A1").End(xlToRight)).Count - 1
    '
    Dim rng As Range
    Dim ColMean As Double, ColStddev As Double
        '
        'calculate parameters for each factor
        For i = 0 To ColNum - 1
            '
            Set rng = ActiveSheet.Range(Range("B2").Offset(0, i), Range("B2").Offset(lastRow, i))
            ColMean = WorksheetFunction.Average(rng)
            ColStddev = WorksheetFunction.StDev_P(rng)
            '
            'Calc & Print z-scores for each security
            Range("B1").Offset(0, i + ColNum).Value = "Z_" & Range("B1").Offset(0, i).Value
            On Error Resume Next
            For j = 0 To lastRow - 2
                If IsEmpty(Range("B2").Offset(j, i).Value) Then
                    Range("B2").Offset(j, i + ColNum).Value = ""
                Else
                    Range("B2").Offset(j, i + ColNum).Value = _
                    (Range("B2").Offset(j, i).Value - ColMean) / ColStddev
                End If
            Next j
            On Error GoTo 0
        Next i
        '
        'Calc aggregate score
        Range("B1").Offset(0, ColNum * 2).Value = factor
        On Error Resume Next
        For i = 0 To lastRow - 2
            Set rng = ActiveSheet.Range(Range("B2").Offset(i, ColNum), Range("B2").Offset(i, ColNum * 2 - 1))
            Range("B2").Offset(i, ColNum * 2).Value = WorksheetFunction.Average(rng)
        Next i
        On Error GoTo 0
        '
End Function
'
Private Function GetSheets() As Variant
    '
    Dim i As Integer
    Dim ws() As Variant
    '
    ThisWorkbook.Activate 'replace with hardcoded WB name
    '
    For i = 0 To Worksheets.Count - 1
        '
        ReDim Preserve ws(i)
        ws(i) = Worksheets(i + 1).Name
    Next i
    '
    GetSheets = ws
    '
End Function
'
Private Function cleanNAs()
'
    Dim rng As Range
    Sheets(ws_name).Activate
    '
    Set rng = ActiveSheet.Range(Range("B1").End(xlDown), Range("B1").End(xlToRight))
    rng.Replace "#N/A", Null, xlWhole, MatchCase:=True
    '
End Function
'
Private Function AddSheetandTickers()
    '
    Dim sheetDestination As Worksheet
    '
    'Add new sheet
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = ws_name & "_" & factor
    '
    'Copy & Paste tickers
    Sheets(ws_name).Range("A1:A" & lastRow).Copy Destination:=Sheets(ws_name & "_" & factor).Range("A1")
    '
    'determine destination
    Select Case factor
        Case "PROF"
             Set sheetDestination = Sheets(ws_name)
        Case "GROW"
            Set sheetDestination = Sheets(ws_name & "_PROF")
        Case "SAFE"
            Set sheetDestination = Sheets(ws_name & "_GROW")
        Case "PAYO"
            Set sheetDestination = Sheets(ws_name & "_SAFE")

    End Select
    '
    'move sheet
    Sheets(ws_name & "_" & factor).Move after:=sheetDestination
    '
End Function
'
Private Function GetTickers(SheetName As String)
    '
    Dim ArraySecLen As Long, i As Long
    '
    ' determine array lengths
    Sheets(SheetName).Activate
    ArraySecLen = Sheets(SheetName).Range(Range("A2"), Range("A2").End(xlDown)).Count
    '
    ReDim s(0 To ArraySecLen - 1)
    '
    For i = 0 To UBound(s)
        '
        s(i) = Sheets(SheetName).Range("A2").Offset(i, 0).Value
    Next i
    '
    GetTickers = s
    '
End Function
'
Private Function fieldCopy(ByRef f As String, ByRef ws_name As String, ByVal lastRow, ByRef j As Integer)
'
    Dim ColNum As Integer
    '
    Sheets(ws_name).Activate
    Cells.Find(f, searchorder:=xlByColumns, searchdirection:=xlNext).Activate
    ColNum = ActiveCell.Column - 1
    '
    ActiveSheet.Range(Range("A1").Offset(0, ColNum), Range("A1").Offset(lastRow, ColNum)).Copy Destination:= _
    Sheets(ws_name & "_" & factor).Range("A1").Offset(0, j + 1)
    '
End Function
'
Private Function fieldArray(f As String, ByRef ws_name As String, ByVal lastRow As Long) As Variant
'
    ' Function returns array of format (i + 1, 1)
    Dim ColNum As Integer
    '
    Sheets(ws_name).Activate
    Cells.Find(f, searchorder:=xlByColumns, searchdirection:=xlNext).Activate
    ColNum = ActiveCell.Column - 1
    '
     fieldArray = Range(Range("A2").Offset(0, ColNum), Range("A2").Offset(lastRow, ColNum)).Value
    '
End Function
'
Private Function fieldarrayHist(ByRef s As Variant, f As String, ByRef ws_name As String, _
    ByVal yearsback As Integer) As Variant
    '
    ' Function returns array of format (i)
    Dim ColNum As Integer, hLastRow As Integer, i As Integer
    Dim hYear As String
    Dim hSheet As String, tick As String
    Dim rng As Range, foundRng As Range
    Dim fldArry() As Variant
    '
    'determine sheet from periods back
    hYear = Left(ws_name, 4) - yearsback
    hSheet = hYear & Right(ws_name, 4)
    '
    Sheets(hSheet).Activate
    hLastRow = Range("A2").End(xlDown).Row
    Cells.Find(f, searchorder:=xlByColumns, searchdirection:=xlNext).Activate
    ColNum = ActiveCell.Column - 1
    '
    Set rng = Range("A2:A" & hLastRow)
    For i = 0 To UBound(s)
        ReDim Preserve fldArry(i)
        Set foundRng = rng.Find(s(i), searchorder:=xlByRows, searchdirection:=xlNext)
        If foundRng Is Nothing Then
            fldArry(i) = Null
        Else
            foundRng.Activate
            fldArry(i) = ActiveCell.Offset(0, ColNum).Value
        End If
    Next i
    '
    fieldarrayHist = fldArry
    '
End Function
'
Public Sub DeleteSheets()
    '
    Dim sh As Worksheet
    Dim suf As String
    '
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    '
    For Each sh In Worksheets
        suf = Right(sh.Name, 4)
        If suf = "PROF" Or suf = "SAFE" Or suf = "PAYO" Or suf = "GROW" Then
            sh.Delete
        End If
    Next
    '
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '
End Sub
