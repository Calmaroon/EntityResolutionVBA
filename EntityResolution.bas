Attribute VB_Name = "modEntityResolution"
'===============================================================================
' Entity resolution for person names - groups records that refer to the same
' person, optionally using date of birth and an ID (SSN, email, employee no.).
'
'   ER_Resolve(Array("John Smith", "John Q Smith", "John Quincy Smith", _
'                    "William Jones", "Bill Jones", "Mary Smith"))
'       = (1, 1, 1, 2, 2, 3)
'
' How it works
'   1. Each name is parsed into first / middle / last / suffix. "Last, First"
'      order, titles (Dr, Mrs), credentials (PhD), accents, punctuation and
'      surname particles (van, de, St) are handled.
'   2. Two records are compared part by part. A part agrees when it is exact,
'      a nickname (Bill = William), a one-letter typo (Damerau-Levenshtein),
'      sounds alike (Double Metaphone, MRA or Editex, with Jaro-Winkler >= 0.8)
'      or, for given names, an initial (Q = Quincy). A part that clearly
'      disagrees (John vs Mary, Eric vs Erica, Jr vs Sr, a different date of
'      birth or ID) is a conflict and the two records can never be the same
'      person.
'   3. Score = average of the first- and last-name scores, plus small bonuses
'      for a matching middle name and date of birth (capped at 1). Pairs at or
'      above Threshold are linked, strongest first, and a link is only made
'      when no record in one group conflicts with any record in the other. So
'      "John Q Smith" and "John R Smith" stay apart even though both match
'      "John Smith".
'
' Requires
'   JaroWinkler.bas, DamerauLevenshtein.bas and Phonetic.bas from TextDistanceVBA
'     (https://github.com/Calmaroon/TextDistanceVBA)
'   Phonetic_Encodings/Encode_DoubleMetaphone.bas from PhoneticVBA
'     (https://github.com/Calmaroon/PhoneticVBA)
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'===============================================================================
Option Explicit
Option Compare Binary

Public Const ER_DEFAULT_THRESHOLD As Double = 0.85

' Part scores
Private Const SCORE_EXACT As Double = 1
Private Const SCORE_NICKNAME As Double = 0.95
Private Const SCORE_TYPO_GIVEN As Double = 0.9
Private Const SCORE_TYPO_LAST As Double = 0.92
Private Const SCORE_SOUNDS_GIVEN As Double = 0.88
Private Const SCORE_SOUNDS_LAST As Double = 0.9
Private Const SCORE_INITIAL As Double = 0.85
Private Const SCORE_MISSING As Double = 0.6        ' first or last name absent on one side
Private Const BONUS_MIDDLE_FULL As Double = 0.03
Private Const BONUS_MIDDLE_INITIAL As Double = 0.01
Private Const BONUS_DOB As Double = 0.05
Private Const MIN_SOUNDS_JW As Double = 0.8
Private Const BLOCK_PASSES As Long = 5
Private Const MIN_EDITEX_JW As Double = 0.85
Private Const MIN_EDITEX As Double = 0.75

' Part comparison results that are not scores
Private Const PART_MISSING As Double = -1
Private Const PART_CONFLICT As Double = 0

Private Type ERRecord
    Blank As Boolean
    Original As String
    FirstName As String
    MiddleName As String        ' first middle token
    MiddleAll As String
    LastName As String
    Suffix As String            ' JR, SR, II, III, IV
    FirstCode As String         ' "DMPRIMARY,DMSECONDARY;MRACODEX"
    MiddleCode As String
    LastCode As String
    FirstNick As String         ' nickname groups, e.g. ",12,40,"
    MiddleNick As String
    DOB As String               ' YYYYMMDD, or "?digits" when not a date
    ID As String
End Type

Private mGroups() As String     ' "|WILLIAM|BILL|BILLY|"
Private mGroupCount As Long
Private mGroupsLoaded As Boolean

'-------------------------------------------------------------------------------
' Public functions
'-------------------------------------------------------------------------------

' Assigns an entity number to every name: records judged to be the same person
' share a number. Numbers start at 1 in order of first appearance; blank names
' give "".
'
'   Names       Range, 1-D or 2-D array, or a single value.
'   DOBs, IDs   Optional, same size as Names. Dates, "YYYY-MM-DD" text, or
'               YYYYMMDD numbers for DOBs; any text for IDs (compared exactly,
'               ignoring case and spaces). Blank values are simply ignored.
'   Threshold   Minimum score (0..1) to link two records. Default 0.85.
'   Blocking    True (default) only compares records whose surnames sound alike,
'               share the first 3 letters or letters 2-4, start with the same
'               sounds as well as the first names, or that share an ID. Much faster on big lists;
'               False compares every pair.
'   ReturnNames True returns, instead of numbers, the fullest name in each
'               group (e.g. "John Quincy Smith" for all three Johns).
'
' The result has the same shape as Names, so in Excel 365 =ER_Resolve(A2:A100)
' spills down a column.
Public Function ER_Resolve(ByVal Names As Variant, _
                           Optional ByVal DOBs As Variant, _
                           Optional ByVal IDs As Variant, _
                           Optional ByVal Threshold As Double = ER_DEFAULT_THRESHOLD, _
                           Optional ByVal Blocking As Boolean = True, _
                           Optional ByVal ReturnNames As Boolean = False) As Variant
    Dim nameList As Variant, dobList As Variant, idList As Variant
    Dim rows As Long, cols As Long, shape As Long, base1 As Long
    Dim n As Long, i As Long, j As Long, k As Long
    Dim recs() As ERRecord

    n = Flatten(Names, nameList, rows, cols, shape, base1)
    If n = 0 Then
        ER_Resolve = Empty
        Exit Function
    End If
    If Not IsMissing(DOBs) Then
        If Flatten(DOBs, dobList, 0, 0, 0, 0) <> n Then
            Err.Raise 5, "ER_Resolve", "DOBs must have the same number of values as Names"
        End If
    End If
    If Not IsMissing(IDs) Then
        If Flatten(IDs, idList, 0, 0, 0, 0) <> n Then
            Err.Raise 5, "ER_Resolve", "IDs must have the same number of values as Names"
        End If
    End If

    ReDim recs(1 To n)
    For i = 1 To n
        If IsMissing(DOBs) And IsMissing(IDs) Then
            BuildRecord recs(i), nameList(i), Empty, Empty
        ElseIf IsMissing(IDs) Then
            BuildRecord recs(i), nameList(i), dobList(i), Empty
        ElseIf IsMissing(DOBs) Then
            BuildRecord recs(i), nameList(i), Empty, idList(i)
        Else
            BuildRecord recs(i), nameList(i), dobList(i), idList(i)
        End If
    Next i

    ' Candidate links
    Dim eA() As Long, eB() As Long, eScore() As Double, edgeCount As Long
    ReDim eA(1 To 16): ReDim eB(1 To 16): ReDim eScore(1 To 16)
    If Blocking Then
        Dim keys() As String, allKeys() As String
        ReDim keys(1 To n): ReDim allKeys(1 To n, 1 To BLOCK_PASSES)
        For k = 1 To BLOCK_PASSES
            For i = 1 To n
                keys(i) = BlockKey(recs(i), k)
                allKeys(i, k) = keys(i)
            Next i
            AddBlockEdges recs, keys, allKeys, k, n, Threshold, eA, eB, eScore, edgeCount
        Next k
    Else
        For i = 1 To n - 1
            For j = i + 1 To n
                AddEdge recs, i, j, Threshold, eA, eB, eScore, edgeCount
            Next j
        Next i
    End If

    ' Link strongest pairs first, never joining groups that contain a conflict
    Dim parent() As Long, nxt() As Long, order() As Long, ra As Long, rb As Long, t As Long
    ReDim parent(1 To n): ReDim nxt(1 To n)
    For i = 1 To n
        parent(i) = i
        nxt(i) = i
    Next i
    If edgeCount > 0 Then
        ReDim order(1 To edgeCount)
        For i = 1 To edgeCount
            order(i) = i
        Next i
        SortEdges order, eScore, edgeCount
        For k = 1 To edgeCount
            ra = FindRoot(parent, eA(order(k)))
            rb = FindRoot(parent, eB(order(k)))
            If ra <> rb Then
                If CanMerge(recs, nxt, ra, rb) Then
                    If ra > rb Then t = ra: ra = rb: rb = t
                    parent(rb) = ra
                    t = nxt(ra): nxt(ra) = nxt(rb): nxt(rb) = t
                End If
            End If
        Next k
    End If

    ' Number the groups in order of first appearance
    Dim label() As Variant, groupId() As Long, nextId As Long, best() As Long, r As Long
    ReDim label(1 To n): ReDim groupId(1 To n): ReDim best(1 To n)
    For i = 1 To n
        If recs(i).Blank Then
            label(i) = ""
        Else
            r = FindRoot(parent, i)
            If groupId(r) = 0 Then
                nextId = nextId + 1
                groupId(r) = nextId
                best(r) = i
            ElseIf Fullness(recs(i)) > Fullness(recs(best(r))) Then
                best(r) = i
            End If
            label(i) = groupId(r)
        End If
    Next i
    If ReturnNames Then
        For i = 1 To n
            If Not recs(i).Blank Then label(i) = recs(best(FindRoot(parent, i))).Original
        Next i
    End If

    ER_Resolve = Unflatten(label, rows, cols, shape, base1)
End Function

' Score from 0 to 1 for two names being the same person. 0 means a conflict
' (they cannot be the same person). Records are linked at >= 0.85 by default.
'   ER_MatchScore("Bill Jones", "William Jones")        = 0.975
'   ER_MatchScore("John Smith", "Mary Smith")   = 0
Public Function ER_MatchScore(ByVal Name1 As Variant, ByVal Name2 As Variant, _
                              Optional ByVal DOB1 As Variant, Optional ByVal DOB2 As Variant, _
                              Optional ByVal ID1 As Variant, Optional ByVal ID2 As Variant) As Double
    Dim a As ERRecord, b As ERRecord, conflict As Boolean, why As String
    BuildPair a, b, Name1, Name2, DOB1, DOB2, ID1, ID2
    ER_MatchScore = ComparePair(a, b, conflict, why, False)
End Function

' True when ER_MatchScore >= Threshold.
Public Function ER_IsMatch(ByVal Name1 As Variant, ByVal Name2 As Variant, _
                           Optional ByVal DOB1 As Variant, Optional ByVal DOB2 As Variant, _
                           Optional ByVal ID1 As Variant, Optional ByVal ID2 As Variant, _
                           Optional ByVal Threshold As Double = ER_DEFAULT_THRESHOLD) As Boolean
    ER_IsMatch = ER_MatchScore(Name1, Name2, DOB1, DOB2, ID1, ID2) >= Threshold
End Function

' The score with the reason for each part, for reviewing or tuning matches:
'   ER_Explain("Bill Jones", "William Jones")
'   = "0.975: first nickname (BILL~WILLIAM); middle -; last exact (JONES)"
Public Function ER_Explain(ByVal Name1 As Variant, ByVal Name2 As Variant, _
                           Optional ByVal DOB1 As Variant, Optional ByVal DOB2 As Variant, _
                           Optional ByVal ID1 As Variant, Optional ByVal ID2 As Variant) As String
    Dim a As ERRecord, b As ERRecord, conflict As Boolean, why As String, score As Double
    BuildPair a, b, Name1, Name2, DOB1, DOB2, ID1, ID2
    score = ComparePair(a, b, conflict, why, True)
    If conflict Then
        ER_Explain = "0 (conflict): " & why
    Else
        ER_Explain = Format$(score, "0.000") & ": " & why
    End If
End Function

' The cleaned-up name, or one part of it: Part = "FIRST", "MIDDLE", "LAST" or
' "SUFFIX". Without Part, gives "FIRST MIDDLE LAST SUFFIX".
'   ER_ParseName("Smith, John Q.")           = "JOHN Q SMITH"
'   ER_ParseName("Dr. Ludwig van Beethoven", "LAST") = "VANBEETHOVEN"
Public Function ER_ParseName(ByVal Name As Variant, Optional ByVal Part As String = "") As String
    Dim r As ERRecord
    BuildRecord r, Name, Empty, Empty
    Select Case UCase$(Trim$(Part))
        Case "FIRST": ER_ParseName = r.FirstName
        Case "MIDDLE": ER_ParseName = r.MiddleAll
        Case "LAST": ER_ParseName = r.LastName
        Case "SUFFIX": ER_ParseName = r.Suffix
        Case "": ER_ParseName = Trim$(Replace(r.FirstName & " " & r.MiddleAll & " " & _
                                              r.LastName & " " & r.Suffix, "  ", " "))
        Case Else
            Err.Raise 5, "ER_ParseName", "Part must be FIRST, MIDDLE, LAST or SUFFIX"
    End Select
End Function

' Adds a group of interchangeable given names to the built-in nickname list,
' separated by | or commas: ER_AddNicknames "Ivan|Vanya"
Public Sub ER_AddNicknames(ByVal Group As String)
    Dim parts() As String, i As Long, s As String
    If Not mGroupsLoaded Then LoadNicknames
    parts = Split(Replace(Group, ",", "|"), "|")
    For i = 0 To UBound(parts)
        If CleanToken(parts(i)) <> "" Then s = s & CleanToken(parts(i)) & "|"
    Next i
    If s = "" Then Exit Sub
    mGroupCount = mGroupCount + 1
    ReDim Preserve mGroups(1 To mGroupCount)
    mGroups(mGroupCount) = "|" & s
End Sub

' Restores the built-in nickname list, dropping groups added with ER_AddNicknames.
Public Sub ER_ResetNicknames()
    mGroupsLoaded = False
    LoadNicknames
End Sub

'-------------------------------------------------------------------------------
' Comparing two records
'-------------------------------------------------------------------------------

' Returns the score; Conflict is set when the records cannot be the same person
' (the score is then 0). Why gets the per-part reasons when Explain is True.
Private Function ComparePair(a As ERRecord, b As ERRecord, ByRef Conflict As Boolean, _
                             ByRef Why As String, ByVal Explain As Boolean) As Double
    Dim f As Double, m As Double, l As Double, score As Double
    Dim fWhy As String, mWhy As String, lWhy As String, dWhy As String
    Dim dobState As Long

    Conflict = False
    If a.Blank Or b.Blank Then
        Conflict = True
        Why = "blank name"
        Exit Function
    End If

    ' An ID settles it either way
    If a.ID <> "" And b.ID <> "" Then
        If a.ID = b.ID Then
            ComparePair = 1
            If Explain Then Why = "id exact"
        Else
            Conflict = True
            If Explain Then Why = "id different (" & a.ID & " vs " & b.ID & ")"
        End If
        Exit Function
    End If

    ' Cheap parts first, so most non-matches stop early
    If a.Suffix <> "" And b.Suffix <> "" And a.Suffix <> b.Suffix Then Conflict = True
    dobState = CompareDOB(a.DOB, b.DOB, dWhy)
    If dobState = 0 Then Conflict = True
    l = LastScore(a, b, lWhy)
    If l = PART_CONFLICT Then Conflict = True
    If Conflict And Not Explain Then Exit Function

    f = GivenScore(a.FirstName, a.FirstCode, a.FirstNick, b.FirstName, b.FirstCode, b.FirstNick, fWhy)
    m = GivenScore(a.MiddleName, a.MiddleCode, a.MiddleNick, b.MiddleName, b.MiddleCode, b.MiddleNick, mWhy)
    ' "J Robert Smith" goes by his middle name: "Robert Smith"
    If f = PART_CONFLICT Then
        If GoesByMiddle(a, b) Or GoesByMiddle(b, a) Then
            f = SCORE_INITIAL
            fWhy = "uses middle name"
            m = PART_MISSING
            mWhy = "-"
        End If
    End If

    If Explain Then
        Why = "first " & fWhy & "; middle " & mWhy & "; last " & lWhy
        If a.Suffix <> "" Or b.Suffix <> "" Then Why = Why & "; suffix " & _
            IIf(a.Suffix = b.Suffix, "exact", IIf(a.Suffix = "" Or b.Suffix = "", "-", "different"))
        If dWhy <> "" Then Why = Why & "; dob " & dWhy
    End If

    If f = PART_CONFLICT Or m = PART_CONFLICT Then Conflict = True
    If Conflict Then Exit Function

    If f = PART_MISSING Then f = SCORE_MISSING
    If l = PART_MISSING Then l = SCORE_MISSING
    score = (f + l) / 2
    If m >= SCORE_NICKNAME Then
        score = score + BONUS_MIDDLE_FULL
    ElseIf m > 0 Then
        score = score + BONUS_MIDDLE_INITIAL
    End If
    If dobState = 2 Then score = score + BONUS_DOB
    If score > 1 Then score = 1
    ComparePair = score
End Function

' First or middle names. Returns a score, PART_MISSING or PART_CONFLICT.
Private Function GivenScore(ByVal n1 As String, ByVal code1 As String, ByVal nick1 As String, _
                            ByVal n2 As String, ByVal code2 As String, ByVal nick2 As String, _
                            ByRef Why As String) As Double
    If n1 = "" Or n2 = "" Then
        GivenScore = PART_MISSING
        Why = "-"
    ElseIf n1 = n2 Then
        GivenScore = SCORE_EXACT
        Why = "exact (" & n1 & ")"
    ElseIf Len(n1) = 1 Or Len(n2) = 1 Then
        If Left$(n1, 1) = Left$(n2, 1) Then
            GivenScore = SCORE_INITIAL
            Why = "initial (" & n1 & "~" & n2 & ")"
        Else
            GivenScore = PART_CONFLICT
            Why = "different (" & n1 & " vs " & n2 & ")"
        End If
    ElseIf SharesGroup(nick1, nick2) Then
        GivenScore = SCORE_NICKNAME
        Why = "nickname (" & n1 & "~" & n2 & ")"
    ElseIf IsGenderVariant(n1, n2) Then
        GivenScore = PART_CONFLICT
        Why = "different (" & n1 & " vs " & n2 & ")"
    ElseIf IsTypo(n1, n2) Then
        GivenScore = SCORE_TYPO_GIVEN
        Why = "spelling (" & n1 & "~" & n2 & ")"
    ElseIf SoundsAlike(n1, code1, n2, code2) Then
        GivenScore = SCORE_SOUNDS_GIVEN
        Why = "sounds alike (" & n1 & "~" & n2 & ")"
    Else
        GivenScore = PART_CONFLICT
        Why = "different (" & n1 & " vs " & n2 & ")"
    End If
End Function

Private Function LastScore(a As ERRecord, b As ERRecord, ByRef Why As String) As Double
    If a.LastName = "" Or b.LastName = "" Then
        LastScore = PART_MISSING
        Why = "-"
    ElseIf a.LastName = b.LastName Then
        LastScore = SCORE_EXACT
        Why = "exact (" & a.LastName & ")"
    ElseIf IsTypo(a.LastName, b.LastName) Then
        LastScore = SCORE_TYPO_LAST
        Why = "spelling (" & a.LastName & "~" & b.LastName & ")"
    ElseIf SoundsAlike(a.LastName, a.LastCode, b.LastName, b.LastCode) Then
        LastScore = SCORE_SOUNDS_LAST
        Why = "sounds alike (" & a.LastName & "~" & b.LastName & ")"
    Else
        LastScore = PART_CONFLICT
        Why = "different (" & a.LastName & " vs " & b.LastName & ")"
    End If
End Function

' a has a first initial and a full middle name that matches b's first name.
Private Function GoesByMiddle(a As ERRecord, b As ERRecord) As Boolean
    Dim why As String
    If Len(a.FirstName) <> 1 Or Len(a.MiddleName) < 2 Or Len(b.FirstName) < 2 Then Exit Function
    GoesByMiddle = GivenScore(a.MiddleName, a.MiddleCode, a.MiddleNick, _
                              b.FirstName, b.FirstCode, b.FirstNick, why) >= SCORE_SOUNDS_GIVEN
End Function

' Eric / Erica, Mario / Maria, Christian / Christina: endings differ and one
' ends in A or O. Look like typos but are usually different people. A trailing
' H is ignored (Hanna / Hannah).
Private Function IsGenderVariant(ByVal s1 As String, ByVal s2 As String) As Boolean
    Dim e1 As String, e2 As String
    If Right$(s1, 1) = "H" And Len(s1) > 2 Then s1 = Left$(s1, Len(s1) - 1)
    If Right$(s2, 1) = "H" And Len(s2) > 2 Then s2 = Left$(s2, Len(s2) - 1)
    e1 = Right$(s1, 1)
    e2 = Right$(s2, 1)
    IsGenderVariant = e1 <> e2 And (e1 = "A" Or e1 = "O" Or e2 = "A" Or e2 = "O")
End Function

' One edit apart (two for long names), ignoring very short names where one edit
' is a different name (Jon / Jan).
Private Function IsTypo(ByVal s1 As String, ByVal s2 As String) As Boolean
    Dim d As Long, shortest As Long
    shortest = Len(s1)
    If Len(s2) < shortest Then shortest = Len(s2)
    If shortest < 4 Or Abs(Len(s1) - Len(s2)) > 2 Then Exit Function
    d = DamerauLevenshtein_Distance(s1, s2)
    IsTypo = d = 1 Or (d = 2 And shortest >= 8)
End Function

' Same Double Metaphone or MRA code, or a close Editex (sound-aware edit
' distance) - and not too different in spelling (Jaro-Winkler).
Private Function SoundsAlike(ByVal s1 As String, ByVal code1 As String, _
                             ByVal s2 As String, ByVal code2 As String) As Boolean
    Dim jw As Double
    jw = JaroWinkler_Similarity(s1, s2)
    If jw < MIN_SOUNDS_JW Then Exit Function
    If CodesOverlap(Split(code1, ";")(0), Split(code2, ";")(0)) Then
        SoundsAlike = True
    ElseIf Split(code1, ";")(1) = Split(code2, ";")(1) Then
        SoundsAlike = True
    ElseIf jw >= MIN_EDITEX_JW Then
        SoundsAlike = Editex_NormalizedSimilarity(s1, s2) >= MIN_EDITEX
    End If
End Function

Private Function CodesOverlap(ByVal code1 As String, ByVal code2 As String) As Boolean
    Dim p1() As String, p2() As String, i As Long, j As Long
    p1 = Split(code1, ",")
    p2 = Split(code2, ",")
    For i = 0 To UBound(p1)
        If p1(i) <> "" Then
            For j = 0 To UBound(p2)
                If p1(i) = p2(j) Then
                    CodesOverlap = True
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

' 2 = same date, 1 = close (one digit or day/month swapped) or unknown,
' 0 = different.
Private Function CompareDOB(ByVal d1 As String, ByVal d2 As String, ByRef Why As String) As Long
    Why = ""
    If d1 = "" Or d2 = "" Then
        CompareDOB = 1
    ElseIf d1 = d2 Then
        CompareDOB = 2
        Why = "exact"
    ElseIf Left$(d1, 1) <> "?" And Left$(d2, 1) <> "?" And _
           (DamerauLevenshtein_Distance(d1, d2) = 1 Or _
            (Left$(d1, 4) = Left$(d2, 4) And Mid$(d1, 5, 2) = Mid$(d2, 7, 2) And Mid$(d1, 7, 2) = Mid$(d2, 5, 2))) Then
        CompareDOB = 1
        Why = "close (" & d1 & "~" & d2 & ")"
    Else
        CompareDOB = 0
        Why = "different (" & d1 & " vs " & d2 & ")"
    End If
End Function

'-------------------------------------------------------------------------------
' Parsing
'-------------------------------------------------------------------------------

Private Sub BuildPair(a As ERRecord, b As ERRecord, ByVal Name1 As Variant, ByVal Name2 As Variant, _
                      Optional ByVal DOB1 As Variant, Optional ByVal DOB2 As Variant, _
                      Optional ByVal ID1 As Variant, Optional ByVal ID2 As Variant)
    If IsMissing(DOB1) Then DOB1 = Empty
    If IsMissing(DOB2) Then DOB2 = Empty
    If IsMissing(ID1) Then ID1 = Empty
    If IsMissing(ID2) Then ID2 = Empty
    BuildRecord a, Name1, DOB1, ID1
    BuildRecord b, Name2, DOB2, ID2
End Sub

Private Sub BuildRecord(r As ERRecord, ByVal Name As Variant, ByVal DOB As Variant, ByVal ID As Variant)
    Dim s As String, tokens() As String, n As Long, i As Long, lastStart As Long
    Dim commaPos As Long, before As String, after As String

    If Not mGroupsLoaded Then LoadNicknames
    s = ToText(Name)
    r.Original = s
    r.FirstName = "": r.MiddleName = "": r.MiddleAll = "": r.LastName = "": r.Suffix = ""
    r.FirstCode = "": r.MiddleCode = "": r.LastCode = "": r.FirstNick = "": r.MiddleNick = ""

    ' Upper case letters, spaces and one comma; O'Brien -> OBRIEN, Smith-Jones -> SMITHJONES
    s = CleanText(s)
    commaPos = InStr(s, ",")
    If commaPos > 0 Then
        before = Trim$(Left$(s, commaPos - 1))
        after = Trim$(Replace(Mid$(s, commaPos + 1), ",", " "))
        If OnlySuffixes(after) Then
            s = before & " " & after          ' "John Smith, Jr."
        Else
            s = after & " " & before          ' "Smith, John Q"
        End If
    End If
    SplitTokens s, tokens, n

    ' Leading titles, trailing suffixes and credentials
    Do While n > 1
        If Not IsTitle(tokens(1)) Then Exit Do
        RemoveToken tokens, n, 1
    Loop
    Do While n > 1
        If IsGenerationSuffix(tokens(n)) Then
            If r.Suffix = "" Then r.Suffix = NormalizeSuffix(tokens(n))
        ElseIf Not IsCredential(tokens(n)) Then
            Exit Do
        End If
        RemoveToken tokens, n, n
    Loop

    r.Blank = (n = 0)
    If n = 1 Then
        r.LastName = tokens(1)
    ElseIf n > 1 Then
        r.FirstName = tokens(1)
        ' Surname particles: van Beethoven, de la Cruz, St John
        lastStart = n
        Do While lastStart > 2
            If Not IsParticle(tokens(lastStart - 1)) Then Exit Do
            lastStart = lastStart - 1
        Loop
        For i = lastStart To n
            r.LastName = r.LastName & tokens(i)
        Next i
        For i = 2 To lastStart - 1
            r.MiddleAll = r.MiddleAll & IIf(i > 2, " ", "") & tokens(i)
        Next i
        If lastStart > 2 Then r.MiddleName = tokens(2)
    End If

    r.FirstCode = PhoneticCode(r.FirstName)
    r.MiddleCode = PhoneticCode(r.MiddleName)
    r.LastCode = PhoneticCode(r.LastName)
    r.FirstNick = NicknameGroups(r.FirstName)
    r.MiddleNick = NicknameGroups(r.MiddleName)
    r.DOB = NormalizeDOB(DOB)
    r.ID = UCase$(Replace(ToText(ID), " ", ""))
End Sub

Private Function ToText(ByVal Value As Variant) As String
    Dim v As Variant
    v = Value                                ' a Range or field gives its value here
    If IsArray(v) Or IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    ToText = Trim$(CStr(v))
End Function

' Accents removed, upper case, only A-Z, spaces and commas.
Private Function CleanText(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    s = UCase$(FoldAccents(s))
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "A" And c <= "Z" Then
            out = out & c
        ElseIf c = "," Then
            out = out & ","
        ElseIf c = "'" Or c = "-" Or c = ChrW(8217) Then
            ' joined: O'Brien, Smith-Jones
        Else
            out = out & " "
        End If
    Next i
    CleanText = out
End Function

Private Function CleanToken(ByVal s As String) As String
    CleanToken = Replace(Replace(CleanText(s), " ", ""), ",", "")
End Function

' Latin-1 and common Latin Extended-A letters to plain letters.
Private Function FoldAccents(ByVal s As String) As String
    Const LATIN1 As String = "AAAAAA_CEEEEIIIIDNOOOOO OUUUUYT_aaaaaa_ceeeeiiiidnooooo ouuuuyty"
    Const EXTENDED_A As String = "AaAaAaCcCcCcCcDdDdEeEeEeEeEeGgGgGgGgHhHhIiIiIiIiIiIiJjKkkLlLlLlLlLlNnNnNnnNnOoOoOoOoRrRrRrSsSsSsSsTtTtTtUuUuUuUuUuUuWwYyYZzZzZzs"
    Dim i As Long, cp As Long, res As String, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        cp = AscW(c)
        If cp < 0 Then cp = cp + 65536
        Select Case cp
            Case 198
                res = res & "AE"
            Case 230
                res = res & "ae"
            Case 223
                res = res & "ss"
            Case 192 To 255
                res = res & Mid$(LATIN1, cp - 191, 1)
            Case 338
                res = res & "OE"
            Case 339
                res = res & "oe"
            Case 256 To 383
                res = res & Mid$(EXTENDED_A, cp - 255, 1)
            Case Else
                res = res & c
        End Select
    Next i
    FoldAccents = res
End Function

' Splits on spaces and commas into tokens(1..n).
Private Sub SplitTokens(ByVal s As String, tokens() As String, ByRef n As Long)
    Dim parts() As String, i As Long
    parts = Split(Trim$(Replace(s, ",", " ")), " ")
    ReDim tokens(1 To UBound(parts) + 2)
    n = 0
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then
            n = n + 1
            tokens(n) = parts(i)
        End If
    Next i
End Sub

Private Sub RemoveToken(tokens() As String, ByRef n As Long, ByVal pos As Long)
    Dim i As Long
    For i = pos To n - 1
        tokens(i) = tokens(i + 1)
    Next i
    n = n - 1
End Sub

Private Function OnlySuffixes(ByVal s As String) As Boolean
    Dim tokens() As String, n As Long, i As Long
    SplitTokens s, tokens, n
    If n = 0 Then Exit Function
    For i = 1 To n
        If Not IsGenerationSuffix(tokens(i)) And Not IsCredential(tokens(i)) Then Exit Function
    Next i
    OnlySuffixes = True
End Function

Private Function InList(ByVal word As String, ByVal list As String) As Boolean
    InList = InStr(" " & list & " ", " " & word & " ") > 0
End Function

Private Function IsTitle(ByVal t As String) As Boolean
    IsTitle = InList(t, "MR MRS MS MISS MX DR PROF REV SIR DAME LADY LORD FR")
End Function

Private Function IsGenerationSuffix(ByVal t As String) As Boolean
    IsGenerationSuffix = InList(t, "JR JNR JUNIOR SR SNR SENIOR II III IV ND RD TH")
End Function

Private Function IsCredential(ByVal t As String) As Boolean
    IsCredential = InList(t, "MD PHD DDS DVM ESQ CPA RN DO JD MBA")
End Function

Private Function NormalizeSuffix(ByVal t As String) As String
    Select Case t
        Case "JNR", "JUNIOR": NormalizeSuffix = "JR"
        Case "SNR", "SENIOR": NormalizeSuffix = "SR"
        Case Else: NormalizeSuffix = t
    End Select
End Function

Private Function IsParticle(ByVal t As String) As Boolean
    IsParticle = InList(t, "VAN VON DE DER DEN DEL DELLA DELA DA DAS DOS DI DU LA LE ST SAINT TER TEN")
End Function

Private Function PhoneticCode(ByVal s As String) As String
    Dim copy As String
    If s = "" Then Exit Function
    copy = s                                 ' DoubleMetaphone changes its argument
    PhoneticCode = DoubleMetaphone(copy) & ";" & MRA_Codex(s)
End Function

' Dates and Excel serials to YYYYMMDD; 8-digit numbers are taken as YYYYMMDD.
' Anything else becomes "?" & its digits and only ever matches exactly.
Private Function NormalizeDOB(ByVal Value As Variant) As String
    Dim v As Variant, s As String, i As Long, digits As String, n As Double
    v = Value                                ' a Range or field gives its value here
    If IsArray(v) Or IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        NormalizeDOB = Format$(v, "yyyymmdd")
        Exit Function
    End If
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then digits = digits & Mid$(s, i, 1)
    Next i
    If Len(digits) = 8 And (Left$(digits, 2) = "19" Or Left$(digits, 2) = "20" Or Left$(digits, 2) = "18") _
       And Len(digits) = Len(s) Then
        NormalizeDOB = digits
    ElseIf IsNumeric(s) Then
        n = CDbl(s)
        If n >= 1 And n < 2958466 Then
            NormalizeDOB = Format$(CDate(n), "yyyymmdd")
        Else
            NormalizeDOB = "?" & digits
        End If
    ElseIf IsDate(s) Then
        NormalizeDOB = Format$(CDate(s), "yyyymmdd")
    Else
        NormalizeDOB = "?" & digits
    End If
End Function

' How much of the name is known, for picking a group's representative name.
Private Function Fullness(r As ERRecord) As Long
    Fullness = Len(r.FirstName) + Len(r.MiddleAll) + Len(r.LastName) + Len(r.Suffix)
End Function

'-------------------------------------------------------------------------------
' Candidate pairs and grouping
'-------------------------------------------------------------------------------

' Pass 1: how the surname sounds. Pass 2: its first 3 letters. Pass 3: the ID.
' Pass 4: the first two sounds of the surname and first sound of the first name
' (Catherine Thompson / Katherine Tomson differ in passes 1 and 2).
' Pass 5: letters 2-4 of the surname and the first initial, for a wrong first
' letter (Ann Chang / Anne Zhang).
Private Function BlockKey(r As ERRecord, ByVal pass As Long) As String
    If r.Blank Then Exit Function
    Select Case pass
        Case 1
            If r.LastName <> "" Then
                BlockKey = Split(Split(r.LastCode, ";")(0) & ",", ",")(0)
                If BlockKey = "" Then BlockKey = r.LastName
            End If
        Case 2
            BlockKey = Left$(r.LastName, 3)
        Case 3
            BlockKey = r.ID
        Case 4
            If r.LastName <> "" And r.FirstName <> "" Then
                BlockKey = Left$(Split(r.LastCode, ";")(0), 2) & "|" & Left$(Split(r.FirstCode, ";")(0), 1)
            End If
        Case 5
            If Len(r.LastName) >= 4 And r.FirstName <> "" Then
                BlockKey = Mid$(r.LastName, 2, 3) & "|" & Left$(r.FirstName, 1)
            End If
    End Select
End Function

' Compares every pair of records sharing a non-empty key, except pairs an
' earlier pass already compared.
Private Sub AddBlockEdges(recs() As ERRecord, keys() As String, allKeys() As String, _
                          ByVal pass As Long, ByVal n As Long, _
                          ByVal Threshold As Double, eA() As Long, eB() As Long, _
                          eScore() As Double, edgeCount As Long)
    Dim idx() As Long, i As Long, j As Long, k As Long, runStart As Long, runEnd As Long
    Dim seen As Boolean
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    SortByKey idx, keys, n
    runStart = 1
    Do While runStart <= n
        runEnd = runStart
        Do While runEnd < n
            If keys(idx(runEnd + 1)) <> keys(idx(runStart)) Then Exit Do
            runEnd = runEnd + 1
        Loop
        If keys(idx(runStart)) <> "" Then
            For i = runStart To runEnd - 1
                For j = i + 1 To runEnd
                    seen = False
                    For k = 1 To pass - 1
                        If allKeys(idx(i), k) <> "" Then
                            If allKeys(idx(i), k) = allKeys(idx(j), k) Then seen = True
                        End If
                    Next k
                    If Not seen Then AddEdge recs, idx(i), idx(j), Threshold, eA, eB, eScore, edgeCount
                Next j
            Next i
        End If
        runStart = runEnd + 1
    Loop
End Sub

Private Sub AddEdge(recs() As ERRecord, ByVal i As Long, ByVal j As Long, ByVal Threshold As Double, _
                    eA() As Long, eB() As Long, eScore() As Double, edgeCount As Long)
    Dim score As Double, conflict As Boolean, why As String
    score = ComparePair(recs(i), recs(j), conflict, why, False)
    If conflict Or score < Threshold Then Exit Sub
    edgeCount = edgeCount + 1
    If edgeCount > UBound(eA) Then
        ReDim Preserve eA(1 To 2 * edgeCount)
        ReDim Preserve eB(1 To 2 * edgeCount)
        ReDim Preserve eScore(1 To 2 * edgeCount)
    End If
    If i < j Then
        eA(edgeCount) = i: eB(edgeCount) = j
    Else
        eA(edgeCount) = j: eB(edgeCount) = i
    End If
    eScore(edgeCount) = score
End Sub

Private Function FindRoot(parent() As Long, ByVal i As Long) As Long
    Dim root As Long, t As Long
    root = i
    Do While parent(root) <> root
        root = parent(root)
    Loop
    Do While parent(i) <> root
        t = parent(i)
        parent(i) = root
        i = t
    Loop
    FindRoot = root
End Function

' No record in group ra conflicts with any record in group rb. Members of a
' group form a ring through nxt().
Private Function CanMerge(recs() As ERRecord, nxt() As Long, ByVal ra As Long, ByVal rb As Long) As Boolean
    Dim i As Long, j As Long, conflict As Boolean, why As String
    i = ra
    Do
        j = rb
        Do
            ComparePair recs(i), recs(j), conflict, why, False
            If conflict Then Exit Function
            j = nxt(j)
        Loop While j <> rb
        i = nxt(i)
    Loop While i <> ra
    CanMerge = True
End Function

' Stable merge sort of idx() by keys(idx()).
Private Sub SortByKey(idx() As Long, keys() As String, ByVal n As Long)
    Dim tmp() As Long
    If n < 2 Then Exit Sub
    ReDim tmp(1 To n)
    MergeSortKey idx, tmp, keys, 1, n
End Sub

Private Sub MergeSortKey(idx() As Long, tmp() As Long, keys() As String, ByVal lo As Long, ByVal hi As Long)
    Dim mid_ As Long, i As Long, j As Long, k As Long
    If lo >= hi Then Exit Sub
    mid_ = (lo + hi) \ 2
    MergeSortKey idx, tmp, keys, lo, mid_
    MergeSortKey idx, tmp, keys, mid_ + 1, hi
    i = lo: j = mid_ + 1: k = lo
    Do While i <= mid_ And j <= hi
        If keys(idx(j)) < keys(idx(i)) Then
            tmp(k) = idx(j): j = j + 1
        Else
            tmp(k) = idx(i): i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid_
        tmp(k) = idx(i): i = i + 1: k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = idx(j): j = j + 1: k = k + 1
    Loop
    For k = lo To hi
        idx(k) = tmp(k)
    Next k
End Sub

' Stable merge sort of order() by score, highest first.
Private Sub SortEdges(order() As Long, score() As Double, ByVal n As Long)
    Dim tmp() As Long
    If n < 2 Then Exit Sub
    ReDim tmp(1 To n)
    MergeSortScore order, tmp, score, 1, n
End Sub

Private Sub MergeSortScore(order() As Long, tmp() As Long, score() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim mid_ As Long, i As Long, j As Long, k As Long
    If lo >= hi Then Exit Sub
    mid_ = (lo + hi) \ 2
    MergeSortScore order, tmp, score, lo, mid_
    MergeSortScore order, tmp, score, mid_ + 1, hi
    i = lo: j = mid_ + 1: k = lo
    Do While i <= mid_ And j <= hi
        If score(order(j)) > score(order(i)) Then
            tmp(k) = order(j): j = j + 1
        Else
            tmp(k) = order(i): i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid_
        tmp(k) = order(i): i = i + 1: k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = order(j): j = j + 1: k = k + 1
    Loop
    For k = lo To hi
        order(k) = tmp(k)
    Next k
End Sub

'-------------------------------------------------------------------------------
' Input and output shapes
'-------------------------------------------------------------------------------

' Copies Values into a 1-based list and returns its length. shape: 0 = single
' value, 1 = 1-D array starting at lowBound, 2 = 2-D array (rows x cols,
' read row by row).
Private Function Flatten(ByVal Values As Variant, ByRef list As Variant, ByRef rows As Long, _
                         ByRef cols As Long, ByRef shape As Long, ByRef lowBound As Long) As Long
    Dim v As Variant, out() As Variant, i As Long, j As Long, n As Long, lb2 As Long, ub2 As Long

    v = Values                               ' a Range gives its .Value array here
    If Not IsArray(v) Then
        ReDim out(1 To 1)
        out(1) = v
        list = out
        rows = 1: cols = 1: shape = 0
        Flatten = 1
        Exit Function
    End If

    On Error Resume Next
    lb2 = LBound(v, 2)
    If Err.Number <> 0 Then
        shape = 1
    Else
        shape = 2
        ub2 = UBound(v, 2)
    End If
    On Error GoTo 0

    If shape = 1 Then
        rows = UBound(v) - LBound(v) + 1: cols = 1: lowBound = LBound(v)
        If rows <= 0 Then Exit Function
        ReDim out(1 To rows)
        For i = LBound(v) To UBound(v)
            n = n + 1
            out(n) = v(i)
        Next i
    Else
        rows = UBound(v, 1) - LBound(v, 1) + 1: cols = ub2 - lb2 + 1
        If rows <= 0 Or cols <= 0 Then Exit Function
        ReDim out(1 To rows * cols)
        For i = LBound(v, 1) To UBound(v, 1)
            For j = lb2 To ub2
                n = n + 1
                out(n) = v(i, j)
            Next j
        Next i
    End If
    list = out
    Flatten = n
End Function

Private Function Unflatten(values() As Variant, ByVal rows As Long, ByVal cols As Long, _
                           ByVal shape As Long, ByVal lowBound As Long) As Variant
    Dim out() As Variant, i As Long, j As Long, n As Long
    Select Case shape
        Case 0
            Unflatten = values(1)
        Case 1
            ReDim out(lowBound To lowBound + rows - 1)
            For i = 1 To rows
                out(lowBound + i - 1) = values(i)
            Next i
            Unflatten = out
        Case Else
            ReDim out(1 To rows, 1 To cols)
            For i = 1 To rows
                For j = 1 To cols
                    n = n + 1
                    out(i, j) = values(n)
                Next j
            Next i
            Unflatten = out
    End Select
End Function

'-------------------------------------------------------------------------------
' Nicknames
'-------------------------------------------------------------------------------

' ",3,17," - the nickname groups that contain the name.
Private Function NicknameGroups(ByVal name As String) As String
    Dim i As Long, s As String
    If name = "" Then Exit Function
    For i = 1 To mGroupCount
        If InStr(mGroups(i), "|" & name & "|") > 0 Then s = s & i & ","
    Next i
    If s <> "" Then NicknameGroups = "," & s
End Function

Private Function SharesGroup(ByVal g1 As String, ByVal g2 As String) As Boolean
    Dim parts() As String, i As Long
    If g1 = "" Or g2 = "" Then Exit Function
    parts = Split(g1, ",")
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then
            If InStr(g2, "," & parts(i) & ",") > 0 Then
                SharesGroup = True
                Exit Function
            End If
        End If
    Next i
End Function

' Common English given names and their nicknames / spellings. A name can be in
' several groups (AL: Albert, Alan, Alexander, Alfred) without those groups
' matching each other.
Private Sub LoadNicknames()
    Dim s As String, parts() As String, i As Long
    s = "ABIGAIL ABBY ABBIE GAIL;ABRAHAM ABE BRAM;ALBERT AL BERT BERTIE;ALAN ALLAN ALLEN AL;"
    s = s & "ALEXANDER ALEX AL ALEC ALEK ALEKSANDR SANDY XANDER LEX SASHA;"
    s = s & "ALEXANDRA ALEX ALEXA LEXI LEXIE SANDRA SANDY SASHA ALI;ALEXIS ALEX LEXI;"
    s = s & "ALFRED AL ALF ALFIE FRED FREDDIE;ALICE ALI ALLIE ALLY;ALLISON ALISON ALLIE ALLY ALI;"
    s = s & "AMANDA MANDY MANDI;ANDREW ANDY DREW ANDRE;ANGELA ANGIE;ANTHONY TONY ANTON ANT ANTONY;"
    s = s & "ANN ANNE ANNA ANNIE NANCY NAN;ARTHUR ART ARTIE;AUGUSTUS GUS AUGIE AUGUST;"
    s = s & "BARBARA BARB BARBIE BABS;BARTHOLOMEW BART;BEATRICE BEA TRIXIE;"
    s = s & "BENJAMIN BEN BENNY BENJI BENNIE;BERNARD BERNIE;BRADLEY BRAD;BRIAN BRYAN;"
    s = s & "CAMERON CAM;CAROLINE CAROLYN CAROL CARRIE;"
    s = s & "CATHERINE CATHARINE KATHERINE KATHARINE KATHRYN CATHY KATHY KATE KATIE KATY KAT KIT KITTY CAT KAY;"
    s = s & "KATHLEEN KATHY KATE KAT;CHARLES CHARLIE CHUCK CHAS CHAZ;CHARLOTTE CHARLIE LOTTIE;"
    s = s & "CHRISTOPHER KRISTOPHER CHRIS KRIS TOPHER KIT;CHRISTIAN CHRIS KRIS;"
    s = s & "CHRISTINA CHRISTINE KRISTINA KRISTINE CHRIS KRIS TINA CHRISSY CHRISTY KRISTY;"
    s = s & "CLIFFORD CLIFF;CYNTHIA CINDY;DANIEL DAN DANNY;DANIELLE DANI DANNI;DAVID DAVE DAVY DAVEY;"
    s = s & "DEBORAH DEBRA DEB DEBBIE DEBBY;DENNIS DENNY DENIS;DIANA DIANE DI;"
    s = s & "DOMINIC DOMINICK DOM;DONALD DON DONNIE DONNY;DOROTHY DOT DOTTIE DORA;DOUGLAS DOUG;"
    s = s & "EDWARD ED EDDIE EDDY NED TED TEDDY;EDMUND ED EDDIE NED;EDWIN ED EDDIE;"
    s = s & "ELEANOR ELLIE NELL NELLIE ELLA NORA;ELIJAH ELI;ELIAS ELI;"
    s = s & "ELIZABETH ELISABETH LIZ LIZZIE LIZA ELIZA BETH BETTY BETSY BETTE LIBBY ELSIE LISA ELLIE BESS BESSIE;"
    s = s & "EMILY EMMY EM MILLIE;EMMA EM EMMY;ERIC ERIK ERICK;EUGENE GENE;EVELYN EVIE EVE;EZEKIEL ZEKE;"
    s = s & "HANNAH HANNA;LEAH LEA;SEAN SHAUN SHAWN;"
    s = s & "FRANCIS FRANK FRAN FRANKIE;FRANCES FRAN FRANNIE FRANKIE FANNY;FRANKLIN FRANK;"
    s = s & "FREDERICK FREDRICK FRED FREDDIE FREDDY FRITZ;GABRIEL GABE;GABRIELLE GABBY GABI;"
    s = s & "GERALD GERRY JERRY;GERARD GERRY;GREGORY GREG GREGG;GWENDOLYN GWEN WENDY;"
    s = s & "HAROLD HAL HARRY;HENRY HANK HARRY HAL;HERBERT HERB BERT;HOWARD HOWIE;HUGH HUGHIE;"
    s = s & "ISAAC IKE;ISABEL ISABELLA ISABELLE IZZY BELLA BELLE ISSY;JACOB JAKE JAKOB;"
    s = s & "JACQUELINE JACKIE JACQUI;JAMES JIM JIMMY JIMMIE JAMIE;JANET JAN;"
    s = s & "JEFFREY JEFFERY GEOFFREY GEOFF JEFF;JENNIFER JEN JENN JENNY JENNI;JEREMIAH JEREMY JERRY;"
    s = s & "JEROME JERRY;JESSICA JESS JESSIE;JOANNA JOANNE JOANN JOAN JO JOANIE;"
    s = s & "JOHN JACK JOHNNY JON JONNY;JONATHAN JONATHON JON JONNY;JOSEPH JOE JOEY JOS;"
    s = s & "JOSEPHINE JO JOSIE JOSEPHINA;JOSHUA JOSH;JUDITH JUDY JUDI JUDE;JULIA JULIE;"
    s = s & "KENNETH KEN KENNY;KIMBERLY KIMBERLEY KIM KIMMY;LAURA LAURIE LORI;"
    s = s & "LAWRENCE LAURENCE LARRY LAURIE;LEONARD LEO LEN LENNY;LOUIS LEWIS LOU LOUIE;"
    s = s & "LOUISE LOUISA LOU;MADELINE MADELEINE MADDIE MADDY;"
    s = s & "MARGARET MAGGIE MEG PEG PEGGY MARGE MARGIE MADGE GRETA GRETCHEN MAISIE MARGO MARGOT DAISY RITA;"
    s = s & "MARTIN MARTY;MARTHA MARTY MATTIE;MARY MOLLY POLLY MAMIE MAE MAY MARIE MARIA;"
    s = s & "MATILDA TILLY TILDA MATTIE;MATTHEW MATHEW MATT MATTY;MAURICE MAURY MO;"
    s = s & "MAXIMILIAN MAX;MAXWELL MAX;MELANIE MEL;MELISSA MEL MISSY LISSA;"
    s = s & "MICHAEL MICHEAL MIKE MIKEY MICK MICKEY MICKY;MICHELLE MICHELE SHELLY SHELLEY;"
    s = s & "MILDRED MILLIE MILLY;MITCHELL MITCH;NATHANIEL NATHANAEL NATHAN NATE NAT;NATHAN NATE;"
    s = s & "NATALIE NAT;NICHOLAS NICOLAS NICK NICKY NIC NIK;NICOLE NIKKI NICKY NIC;NORMAN NORM;"
    s = s & "OLIVER OLLIE OLLY;OLIVIA LIV LIVVY OLLIE;PAMELA PAM PAMMY;"
    s = s & "PATRICIA PAT PATTY PATTI TRICIA TRISH TRISHA PATSY;PATRICK PAT PADDY;PENELOPE PENNY;"
    s = s & "PETER PETE;PHILIP PHILLIP FILIP PHIL PIP;PRISCILLA PRISSY CILLA;PRUDENCE PRU PRUE;"
    s = s & "RANDOLPH RANDY;RANDALL RANDY;RAYMOND RAY;REBECCA REBEKAH BECKY BECCA BECK;"
    s = s & "REGINALD REG REGGIE;RICHARD RICH RICK RICKY RICHIE DICK;"
    s = s & "ROBERT ROB ROBBIE ROBBY BOB BOBBY BOBBIE BERT;ROBERTA BOBBIE ROBBIE BERTA;"
    s = s & "RODNEY ROD RODDY;RODERICK ROD RODDY;ROGER ROG;RONALD RON RONNIE RONNY;"
    s = s & "ROSEMARY ROSEMARIE ROSE ROSIE;RUSSELL RUSS RUSTY;SALVATORE SAL;"
    s = s & "SAMANTHA SAM SAMMY SAMMIE;SAMUEL SAM SAMMY;SANDRA SANDY SANDI;SARAH SARA SALLY SADIE;"
    s = s & "SIDNEY SYDNEY SID;STANLEY STAN;STEPHEN STEVEN STEPHAN STEFAN STEVE STEVIE;"
    s = s & "STEPHANIE STEFANIE STEPH STEFFI;STUART STEWART STU;"
    s = s & "SUSAN SUZAN SUE SUSIE SUZY SUZIE SUZANNE SUSANNA SUSANNAH SUSANNE;"
    s = s & "TERESA THERESA TERRY TERRI TESS TESSA;TERRENCE TERENCE TERRY;THADDEUS TAD THAD;"
    s = s & "THEODORE TED TEDDY THEO;THOMAS TOM TOMMY THOM;TIMOTHY TIM TIMMY;TOBIAS TOBY;"
    s = s & "VALERIE VAL;VICTORIA VICKY VICKI VICKIE TORI TORY VIC;VICTOR VIC;"
    s = s & "VINCENT VINCE VINNY VINNIE VIN;VIRGINIA GINNY GINGER;WALTER WALT WALLY;"
    s = s & "WILLIAM BILL BILLY BILLIE WILL WILLIE WILLY LIAM;WINIFRED WINNIE FREDA;"
    s = s & "ZACHARY ZACHARIAH ZACH ZACK ZAK"

    parts = Split(s, ";")
    mGroupCount = UBound(parts) + 1
    ReDim mGroups(1 To mGroupCount)
    For i = 0 To UBound(parts)
        mGroups(i + 1) = "|" & Replace(parts(i), " ", "|") & "|"
    Next i
    mGroupsLoaded = True
End Sub
