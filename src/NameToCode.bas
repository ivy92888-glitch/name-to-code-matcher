Attribute VB_Name = "NameToCode"
' ============================================================
' name-to-code-matcher
' ------------------------------------------------------------
' Generic, domain-agnostic VBA macro that maps a free-text
' item name (plus an optional attribute such as material) to a
' standardized code, by looking the name up against one or more
' configurable reference tables.
'
' This is an INDEPENDENT, clean-room reimplementation of a
' generic matching pattern. It contains no business-specific
' logic, no proprietary data, and no hard-coded domain codes.
'
' The engine performs:
'   1. Word-boundary-aware substring matching (so "light" does
'      not falsely match inside "lightning").
'   2. Attribute-category filtering (matches are only accepted
'      when the item and reference share the same attribute
'      category, e.g. both are "Metal").
'   3. Multi-tier fallback (tries reference table 1, then 2,
'      then 3; first hit wins).
'   4. A weighted resolver that, when several codes match, picks
'      the best one using a frequency (0.3) + recency (0.7) score.
'
' All sheet names, columns, tier order, weights and the color
' scheme are configurable constants at the top of this module.
' ============================================================

Option Explicit
Option Compare Text

' ---- Source sheet -----------------------------------------
Private Const SRC_SHEET        As String = ""      ' "" = active sheet
Private Const NAME_COL         As String = "C"     ' item name column
Private Const ATTR_COL         As String = "Y"     ' attribute column ("" = none)
Private Const CODE_COL         As String = "E"     ' output code column
Private Const SOURCE_COL       As String = "F"     ' match-source column ("" = none)
Private Const HEADER_ROW       As Long = 3         ' row holding the name header
Private Const DATA_START_ROW   As Long = 4         ' first data row
Private Const NAME_HEADER_TEXT As String = "Item name"  ' header text to expect; "" = skip

' ---- Tiers (reference tables) -----------------------------
' Format of each tier:  SheetName|NameCol|CodeCol|AttrCol
' Tried in order; the first tier that returns a code wins.
Private Const TIERS As String = _
    "RefPrimary|A|B|C,RefSecondary|A|C|H,RefTertiary|A|B|D"

' ---- Attribute taxonomy -----------------------------------
Private Const CATEGORY_SHEET As String = "Categories"  ' keyword -> category sheet
Private Const ATTR_FILTER_ON As Boolean = True         ' filter by attribute category

' ---- Scoring weights --------------------------------------
Private Const W_COUNT   As Double = 0.3   ' weight of match frequency
Private Const W_RECENCY As Double = 0.7   ' weight of row recency (lower = newer)

' ---- Visual feedback --------------------------------------
Private Const COLOR_ON       As Boolean = True
Private Const TIER_COLORS    As String = "198,224,180|255,242,204|255,183,183"
Private Const ERROR_COLOR_R As Long = 255
Private Const ERROR_COLOR_G As Long = 0
Private Const ERROR_COLOR_B As Long = 0
Private Const NO_MATCH_VALUE As String = "N/A"

' ---- Runtime state ----------------------------------------
Private mSuccess As Long
Private mCategories As Object
Private mCategoriesReady As Boolean

' ============================================================
' Entry point
' ============================================================
Public Sub MatchCodes()
    Dim t0 As Double
    t0 = Timer

    Dim srcWB As Workbook, srcWs As Worksheet
    Dim tiers As Collection
    Dim lastRow As Long, processed As Long

    On Error GoTo ErrHandler
    Application.StatusBar = "Initializing matcher..."
    DoEvents

    Set srcWB = ActiveWorkbook
    If SRC_SHEET = "" Then
        Set srcWs = srcWB.ActiveSheet
    Else
        On Error Resume Next
        Set srcWs = srcWB.Sheets(SRC_SHEET)
        On Error GoTo 0
        If srcWs Is Nothing Then
            MsgBox "Source sheet '" & SRC_SHEET & "' not found.", vbCritical, "Config error"
            Exit Sub
        End If
    End If

    If NAME_HEADER_TEXT <> "" Then
        If Trim(srcWs.Cells(HEADER_ROW, NAME_COL).Value) = "" Then
            MsgBox "Header cell " & NAME_COL & HEADER_ROW & " is empty." & vbCrLf & _
                   "Expected header: " & NAME_HEADER_TEXT, vbExclamation, "Format check"
            Exit Sub
        End If
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, NAME_COL).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "No data found in column " & NAME_COL & " from row " & DATA_START_ROW & ".", _
               vbInformation, "Data check"
        Exit Sub
    End If

    Set tiers = GetTierList()
    If Not CheckTierSheets(tiers) Then Exit Sub
    Call EnsureCategories

    mSuccess = 0
    Application.StatusBar = "Matching..."
    DoEvents

    Dim i As Long
    For i = DATA_START_ROW To lastRow
        Dim nm As String
        nm = Trim(srcWs.Cells(i, NAME_COL).Value)
        If nm = "" Then
            srcWs.Cells(i, CODE_COL).Value = ""
            If SOURCE_COL <> "" Then srcWs.Cells(i, SOURCE_COL).Value = ""
            GoTo ContinueLoop
        End If

        Dim pct As Double
        pct = (i - DATA_START_ROW + 1) / (lastRow - DATA_START_ROW + 1) * 100
        Application.StatusBar = "Row " & i & "/" & lastRow & " (" & Format(pct, "0.0") & "%)"
        DoEvents

        Call ProcessRow(srcWs, i, tiers)
        processed = processed + 1
ContinueLoop:
    Next i

    Application.StatusBar = False
    Call ShowDone(lastRow - DATA_START_ROW + 1, processed, mSuccess, Timer - t0)
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Matching failed"
End Sub

' ============================================================
' Per-row processing
' ============================================================
Private Sub ProcessRow(ByVal ws As Worksheet, ByVal row As Long, ByVal tiers As Collection)
    On Error GoTo RowError

    Dim nm As String, attr As String
    nm = Trim(ws.Cells(row, NAME_COL).Value)
    If ATTR_COL = "" Then attr = "" Else attr = Trim(ws.Cells(row, ATTR_COL).Value)
    If nm = "" Then Exit Sub

    Dim idx As Long
    idx = 0
    Dim t As Variant
    For Each t In tiers
        Dim code As String, mLen As Long
        code = MatchFromTier(nm, attr, ThisWorkbook.Sheets(t("Sheet")), _
                             t("Name"), t("Code"), t("Attr"), mLen)
        If code <> "" Then
            Call WriteResult(ws, row, code, "Tier " & (idx + 1) & ": " & t("Sheet"), idx)
            mSuccess = mSuccess + 1
            Exit Sub
        End If
        idx = idx + 1
    Next t

    ws.Cells(row, CODE_COL).Value = NO_MATCH_VALUE
    If COLOR_ON Then ws.Cells(row, CODE_COL).Interior.ColorIndex = xlNone
    If SOURCE_COL <> "" Then ws.Cells(row, SOURCE_COL).Value = "No match"
    Exit Sub

RowError:
    ws.Cells(row, CODE_COL).Value = "ERROR"
    If COLOR_ON Then ws.Cells(row, CODE_COL).Interior.Color = RGB(ERROR_COLOR_R, ERROR_COLOR_G, ERROR_COLOR_B)
    If SOURCE_COL <> "" Then ws.Cells(row, SOURCE_COL).Value = "Error"
End Sub

' ============================================================
' Core matcher against a single reference table
' ============================================================
Private Function MatchFromTier(ByVal itemName As String, ByVal attr As String, _
                               ByVal ws As Worksheet, ByVal nameCol As String, _
                               ByVal codeCol As String, ByVal attrCol As String, _
                               ByRef matchedLen As Long) As String
    Dim lastRow As Long, i As Long
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    If lastRow < 2 Or Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    Dim exact As Object, contain As Object
    Set exact = CreateObject("Scripting.Dictionary")
    Set contain = CreateObject("Scripting.Dictionary")
    matchedLen = 0

    Dim dbName As String, dbCode As String, dbAttr As String
    For i = lastRow To 2 Step -1
        dbName = Trim(ws.Cells(i, nameCol).Value)
        dbCode = Trim(ws.Cells(i, codeCol).Value)
        If attrCol <> "" Then dbAttr = Trim(ws.Cells(i, attrCol).Value) Else dbAttr = ""
        If dbName = "" Or dbCode = "" Then GoTo NextRow

        If ATTR_FILTER_ON And attr <> "" And dbAttr <> "" Then
            If Not IsSameAttributeCategory(attr, dbAttr) Then GoTo NextRow
        End If

        Dim processed As String
        processed = dbName
        Dim dash As Long
        dash = InStr(processed, "-")
        If dash > 0 Then processed = Trim(Left(processed, dash - 1))

        Dim parts As Variant
        parts = Split(processed, "/")

        Dim isExact As Boolean, isContain As Boolean, mLen As Long
        isExact = False: isContain = False: mLen = 0
        Dim p As Variant
        For Each p In parts
            Dim partName As String
            partName = Trim(p)
            If partName = "" Then GoTo SkipPart

            If partName = itemName Then
                isExact = True: mLen = Len(itemName): Exit For
            End If

            If Not isExact Then
                Dim bl As Long
                If BoundaryMatch(partName, itemName, bl) Then
                    isContain = True
                    If bl > mLen Then mLen = bl
                End If
            End If
SkipPart:
        Next p

        If Not isExact And Not isContain Then GoTo NextRow

        If isExact Then
            If Not exact.Exists(dbCode) Then
                Dim e As Object: Set e = CreateObject("Scripting.Dictionary")
                e("code") = dbCode: e("count") = 0: e("latestRow") = i: e("matchLen") = mLen
                exact.Add dbCode, e
            End If
            exact(dbCode)("count") = exact(dbCode)("count") + 1
            If i > exact(dbCode)("latestRow") Then exact(dbCode)("latestRow") = i
            If mLen > exact(dbCode)("matchLen") Then exact(dbCode)("matchLen") = mLen
        Else
            If Not contain.Exists(dbCode) Then
                Dim c As Object: Set c = CreateObject("Scripting.Dictionary")
                c("code") = dbCode: c("count") = 1: c("latestRow") = i: c("matchLen") = mLen
                contain.Add dbCode, c
            Else
                contain(dbCode)("count") = contain(dbCode)("count") + 1
                If mLen > contain(dbCode)("matchLen") Then contain(dbCode)("matchLen") = mLen
                If i > contain(dbCode)("latestRow") Then contain(dbCode)("latestRow") = i
            End If
        End If
NextRow:
    Next i

    If exact.Count > 0 Then
        MatchFromTier = BestCode(exact, lastRow)
        Exit Function
    End If

    If contain.Count > 0 Then
        Dim maxLen As Long, k As Variant
        For Each k In contain.Keys
            If contain(k)("matchLen") > maxLen Then maxLen = contain(k)("matchLen")
        Next k
        Dim filt As Object: Set filt = CreateObject("Scripting.Dictionary")
        For Each k In contain.Keys
            If contain(k)("matchLen") = maxLen Then filt.Add k, contain(k)
        Next k
        MatchFromTier = BestCode(filt, lastRow)
        Exit Function
    End If
End Function

' ============================================================
' Weighted best-code resolver
' ============================================================
Private Function BestCode(ByVal dict As Object, ByVal lastRow As Long) As String
    Dim maxCount As Long, k As Variant
    maxCount = 0
    For Each k In dict.Keys
        If dict(k)("count") > maxCount Then maxCount = dict(k)("count")
    Next k
    If maxCount = 0 Then maxCount = 1

    Dim maxScore As Double, best As String
    maxScore = -1: best = ""
    For Each k In dict.Keys
        Dim cs As Double, rs As Double, ts As Double
        cs = (dict(k)("count") / maxCount) * 100
        If lastRow > 2 Then rs = ((dict(k)("latestRow") - 2) / (lastRow - 2)) * 100 Else rs = 0
        ts = cs * W_COUNT + rs * W_RECENCY
        If ts > maxScore Then maxScore = ts: best = k
    Next k
    BestCode = best
End Function

' ============================================================
' Word-boundary-aware substring match
' Returns True when itemName appears inside partName (partName
' strictly longer) and is surrounded by non-word characters or
' string edges, so "light" won't match "lightning".
' ============================================================
Private Function BoundaryMatch(ByVal partName As String, ByVal itemName As String, _
                               ByRef matchLen As Long) As Boolean
    matchLen = 0
    If Len(partName) <= Len(itemName) Then Exit Function
    Dim pos As Long
    pos = InStr(1, partName, itemName, vbTextCompare)
    If pos = 0 Then Exit Function

    Dim beforeOk As Boolean, afterOk As Boolean
    If pos = 1 Then
        beforeOk = True
    Else
        beforeOk = Not IsWordChar(Mid(partName, pos - 1, 1))
    End If
    If pos + Len(itemName) - 1 = Len(partName) Then
        afterOk = True
    Else
        afterOk = Not IsWordChar(Mid(partName, pos + Len(itemName), 1))
    End If

    If beforeOk And afterOk Then
        BoundaryMatch = True
        matchLen = Len(partName)
    End If
End Function

' A "word character" is a letter, digit, or CJK ideograph.
Private Function IsWordChar(ByVal ch As String) As Boolean
    Dim cc As Integer, uc As Long
    cc = Asc(ch)
    If (cc >= 65 And cc <= 90) Or (cc >= 97 And cc <= 122) Then IsWordChar = True: Exit Function
    If cc >= 48 And cc <= 57 Then IsWordChar = True: Exit Function
    uc = AscW(ch)
    If uc >= 19968 And uc <= 40959 Then IsWordChar = True: Exit Function
    If uc >= 13312 And uc <= 19903 Then IsWordChar = True: Exit Function
    IsWordChar = False
End Function

' ============================================================
' Attribute category filtering
' ============================================================
Private Sub EnsureCategories()
    If mCategoriesReady Then Exit Sub
    Set mCategories = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CATEGORY_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Dim lastRow As Long, i As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            Dim kw As String, cat As String
            kw = Trim(ws.Cells(i, 1).Value)
            cat = Trim(ws.Cells(i, 2).Value)
            If kw <> "" And cat <> "" Then mCategories(kw) = cat
        Next i
    Else
        Call LoadDefaultCategories
    End If
    mCategoriesReady = True
End Sub

Private Sub LoadDefaultCategories()
    ' Sample default taxonomy (physical materials). Replace via the
    ' Categories sheet, or edit this list. Domain-neutral example only.
    Dim rules As Variant
    rules = Array( _
        "plastic,resin,silicone,rubber,acrylic,abs,polyethylene,polyester,polyurethane,pvc,nylon,polycarbonate,pet,pu,tpe,eva,pp,pe,pc", "Plastic / Rubber", _
        "metal,steel,stainless,aluminum,aluminium,copper,brass,alloy,iron,zinc,titanium", "Metal", _
        "wood,bamboo,rattan,veneer,plywood,particle board,mdf", "Wood", _
        "paper,cardboard,corrugated,kraft", "Paper", _
        "fabric,cotton,textile,leather,felt,linen,velvet", "Textile", _
        "glass,tempered glass,quartz,crystal", "Glass", _
        "ceramic,stone,marble,granite,mineral,clay,porcelain", "Mineral", _
        "composite,carbon fiber,fiberglass", "Composite")

    Dim i As Long
    For i = LBound(rules) To UBound(rules) Step 2
        Dim keys As Variant
        keys = Split(rules(i), ",")
        Dim j As Long
        For j = LBound(keys) To UBound(keys)
            Dim k As String
            k = Trim(keys(j))
            If k <> "" Then mCategories(k) = rules(i + 1)
        Next j
    Next i
End Sub

Private Function GetAttributeCategory(ByVal attr As String) As String
    If attr = "" Then GetAttributeCategory = "": Exit Function
    Call EnsureCategories
    Dim low As String, key As Variant
    low = LCase(attr)
    For Each key In mCategories.Keys
        If InStr(1, low, LCase(key), vbTextCompare) > 0 Then
            GetAttributeCategory = mCategories(key): Exit Function
        End If
    Next key
    GetAttributeCategory = ""
End Function

Private Function IsSameAttributeCategory(ByVal a As String, ByVal b As String) As Boolean
    If a = "" Or b = "" Then IsSameAttributeCategory = True: Exit Function
    Dim ca As String, cb As String
    ca = GetAttributeCategory(a)
    cb = GetAttributeCategory(b)
    If ca = "" Or cb = "" Then
        IsSameAttributeCategory = True
    Else
        IsSameAttributeCategory = (ca = cb)
    End If
End Function

' ============================================================
' Output + formatting
' ============================================================
Private Sub WriteResult(ByVal ws As Worksheet, ByVal row As Long, ByVal code As String, _
                        ByVal sourceLabel As String, ByVal tierIndex As Long)
    ws.Cells(row, CODE_COL).Value = code
    If SOURCE_COL <> "" Then ws.Cells(row, SOURCE_COL).Value = sourceLabel
    If COLOR_ON Then
        Dim cols As Variant
        cols = Split(TIER_COLORS, "|")
        If tierIndex >= LBound(cols) And tierIndex <= UBound(cols) Then
            Dim rgb() As String
            rgb = Split(cols(tierIndex), ",")
            If UBound(rgb) = 2 Then
                ws.Cells(row, CODE_COL).Interior.Color = _
                    RGB(CInt(rgb(0)), CInt(rgb(1)), CInt(rgb(2)))
            End If
        End If
    End If
End Sub

' ============================================================
' Configuration helpers
' ============================================================
Private Function GetTierList() As Collection
    Dim raw As Variant
    raw = Split(TIERS, ",")
    Dim c As New Collection
    Dim i As Long
    For i = LBound(raw) To UBound(raw)
        Dim parts As Variant
        parts = Split(raw(i), "|")
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d("Sheet") = Trim(parts(0))
        d("Name") = Trim(parts(1))
        d("Code") = Trim(parts(2))
        If UBound(parts) >= 3 Then d("Attr") = Trim(parts(3)) Else d("Attr") = ""
        c.Add d
    Next i
    Set GetTierList = c
End Function

Private Function CheckTierSheets(ByVal tiers As Collection) As Boolean
    Dim missing As String
    Dim t As Variant
    For Each t In tiers
        Dim ok As Boolean: ok = False
        On Error Resume Next
        ok = Not (ThisWorkbook.Sheets(t("Sheet")) Is Nothing)
        On Error GoTo 0
        If Not ok Then missing = missing & IIf(missing = "", "", ", ") & t("Sheet")
    Next t
    If missing <> "" Then
        MsgBox "Database workbook (the one holding this macro) is missing sheet(s): " & _
               missing & vbCrLf & "Provide them before running.", vbCritical, "Database error"
        CheckTierSheets = False
    Else
        CheckTierSheets = True
    End If
End Function

' ============================================================
' Completion message
' ============================================================
Private Sub ShowDone(ByVal total As Long, ByVal processed As Long, _
                     ByVal success As Long, ByVal dur As Double)
    Dim s As String
    s = "Matching complete!" & vbCrLf & _
        "------------------------------" & vbCrLf & _
        "Total rows:    " & total & vbCrLf & _
        "Processed:     " & processed & vbCrLf & _
        "Matched:       " & success & vbCrLf & _
        "------------------------------" & vbCrLf & _
        "Elapsed:       " & Format(dur, "0.00") & " s"
    MsgBox s, vbInformation, "Done"
End Sub
