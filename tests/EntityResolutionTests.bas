Attribute VB_Name = "modEntityResolutionTests"
' Tests for EntityResolution.bas.
' Needs the same TextDistanceVBA and PhoneticVBA modules as EntityResolution.bas.
' Run Test_EntityResolution and check the Immediate window (Ctrl+G).
Option Explicit
Option Compare Binary

Private mPass As Long, mFail As Long

Public Sub Test_EntityResolution()
    mPass = 0: mFail = 0
    ParseTests
    ScoreTests
    ResolveTests
    Debug.Print "EntityResolution: " & mPass & " passed, " & mFail & " failed"
End Sub

Private Sub ParseTests()
    CheckS "Parse plain", ER_ParseName("John Quincy Smith"), "JOHN QUINCY SMITH"
    CheckS "Parse last, first", ER_ParseName("Smith, John Q."), "JOHN Q SMITH"
    CheckS "Parse suffix after comma", ER_ParseName("John Smith, Jr."), "JOHN SMITH JR"
    CheckS "Parse title and credential", ER_ParseName("Dr. Jane Smith PhD"), "JANE SMITH"
    CheckS "Parse particle", ER_ParseName("Ludwig van Beethoven", "LAST"), "VANBEETHOVEN"
    CheckS "Parse two particles", ER_ParseName("Juan de la Cruz", "LAST"), "DELACRUZ"
    CheckS "Parse apostrophe", ER_ParseName("Conan O'Brien", "LAST"), "OBRIEN"
    CheckS "Parse hyphen", ER_ParseName("Mary Smith-Jones", "LAST"), "SMITHJONES"
    CheckS "Parse accents", ER_ParseName("Jos" & ChrW(233) & " Mu" & ChrW(241) & "oz"), "JOSE MUNOZ"
    CheckS "Parse middle", ER_ParseName("Mary Ann Beth Smith", "MIDDLE"), "ANN BETH"
    CheckS "Parse suffix", ER_ParseName("John Smith Senior", "SUFFIX"), "SR"
    CheckS "Parse single", ER_ParseName("Madonna", "LAST"), "MADONNA"
    CheckS "Parse blank", ER_ParseName("   "), ""
End Sub

Private Sub ScoreTests()
    CheckD "Exact", ER_MatchScore("John Smith", "john smith"), 1
    CheckD "Middle initial", ER_MatchScore("John Q Smith", "John Quincy Smith"), 1
    CheckD "Middle missing", ER_MatchScore("John Smith", "John Quincy Smith"), 1
    CheckD "Nickname", ER_MatchScore("Bill Jones", "William Jones"), 0.975
    CheckD "Different first", ER_MatchScore("John Smith", "Mary Smith"), 0
    CheckD "Different middle", ER_MatchScore("John Q Smith", "John R Smith"), 0
    CheckD "First initial", ER_MatchScore("J Smith", "John Smith"), 0.925
    CheckD "Typo in surname", ER_MatchScore("John Johnson", "John Jonhson"), 0.96
    CheckD "Sounds alike surname", ER_MatchScore("Sarah Thompson", "Sarah Tomson"), 0.95
    CheckD "Spelling given", ER_MatchScore("Katharyne Lee", "Katherine Lee"), 0.95
    CheckD "Sounds alike surname 2", ER_MatchScore("Marc Nielsen", "Mark Nilsson"), 0.9
    CheckD "Sounds alike Editex", ER_MatchScore("Ann Zhang", "Ann Chang"), 0.96
    CheckD "Gender variant", ER_MatchScore("Eric Lee", "Erica Lee"), 0
    CheckD "Gender variant typo", ER_MatchScore("Christian Lee", "Christina Lee"), 0
    CheckD "Mario Maria", ER_MatchScore("Mario Lee", "Maria Lee"), 0
    CheckD "Trailing H", ER_MatchScore("Leah Lee", "Lea Lee"), 0.975
    CheckD "Short surnames differ", ER_MatchScore("Michael Wu", "Michael Wo"), 0
    CheckD "Prefix lookalikes", ER_MatchScore("Christopher Lee", "Christina Lee"), 0
    CheckD "Jr vs Sr", ER_MatchScore("John Smith Jr", "John Smith Sr"), 0
    CheckD "Jr vs none", ER_MatchScore("John Smith Jr", "John Smith"), 1
    CheckD "Goes by middle", ER_MatchScore("J Robert Smith", "Robert Smith"), 0.925
    CheckD "Last, First", ER_MatchScore("Smith, John", "John Smith"), 1
    CheckD "DOB same", ER_MatchScore("J Smith", "John Smith", "1985-03-04", DateSerial(1985, 3, 4)), 0.975
    CheckD "DOB different", ER_MatchScore("John Smith", "John Smith", "1985-03-04", "1990-07-21"), 0
    CheckD "DOB day/month swap", ER_MatchScore("John Smith", "John Smith", "1985-03-04", "1985-04-03"), 1
    CheckD "DOB one digit", ER_MatchScore("John Smith", "John Smith", 19850304, "1985-03-05"), 1
    CheckD "DOB one side", ER_MatchScore("John Smith", "John Smith", "1985-03-04"), 1
    CheckD "ID same", ER_MatchScore("Jane Smith", "Jane Jones", , , "A-1", "a-1"), 1
    CheckD "ID different", ER_MatchScore("Jane Smith", "Jane Smith", , , "A1", "B2"), 0
    CheckD "Surname only", ER_MatchScore("Smith", "John Smith"), 0.8
    CheckB "IsMatch", ER_IsMatch("Bill Jones", "William Jones"), True
    CheckB "IsMatch threshold", ER_IsMatch("J Smith", "John Smith", , , , , 0.95), False
    CheckS "Explain", ER_Explain("Bill Jones", "William Jones"), _
        "0.975: first nickname (BILL~WILLIAM); middle -; last exact (JONES)"
    CheckS "Explain conflict", ER_Explain("John Smith", "Mary Smith"), _
        "0 (conflict): first different (JOHN vs MARY); middle -; last exact (SMITH)"
End Sub

Private Sub ResolveTests()
    Dim r As Variant, names As Variant, dobs As Variant, grid(1 To 3, 1 To 1) As Variant

    names = Array("John Smith", "John Q Smith", "John Quincy Smith", "William Jones", "Bill Jones", "Mary Smith")
    r = ER_Resolve(names)
    CheckS "Resolve example", JoinResult(r), "1,1,1,2,2,3"
    r = ER_Resolve(names, , , , False)
    CheckS "Resolve without blocking", JoinResult(r), "1,1,1,2,2,3"
    r = ER_Resolve(names, , , , , True)
    CheckS "Resolve names", JoinResult(r), _
        "John Quincy Smith,John Quincy Smith,John Quincy Smith,William Jones,William Jones,Mary Smith"

    ' John Smith matches both, but Q and R conflict so they stay apart
    r = ER_Resolve(Array("John Q Smith", "John Smith", "John R Smith"))
    CheckS "Resolve no chaining", JoinResult(r), "1,1,2"

    ' Same name, different dates of birth: two people
    names = Array("John Smith", "Jon Smith", "John Smith", "Johnny Smith")
    dobs = Array("1980-01-01", "", "1975-06-30", "1975-06-30")
    r = ER_Resolve(names, dobs)
    CheckS "Resolve with DOB", JoinResult(r), "1,1,2,2"

    ' Blanks get no number; 2-D input gives 2-D output
    grid(1, 1) = "Bill Jones": grid(2, 1) = Empty: grid(3, 1) = "William Jones"
    r = ER_Resolve(grid)
    CheckL "Resolve 2-D rows", UBound(r, 1), 3
    CheckS "Resolve 2-D", r(1, 1) & "," & r(2, 1) & "," & r(3, 1), "1,,1"

    CheckS "Resolve single", CStr(ER_Resolve("John Smith")), "1"

    ' Nicknames can be extended
    CheckD "Unknown nickname", ER_MatchScore("Vanya Petrov", "Ivan Petrov"), 0
    ER_AddNicknames "Ivan|Vanya"
    CheckD "Added nickname", ER_MatchScore("Vanya Petrov", "Ivan Petrov"), 0.975
    ER_ResetNicknames
    CheckD "Reset nicknames", ER_MatchScore("Vanya Petrov", "Ivan Petrov"), 0
End Sub

Private Function JoinResult(ByVal r As Variant) As String
    Dim i As Long, s As String
    For i = LBound(r) To UBound(r)
        s = s & IIf(i > LBound(r), ",", "") & r(i)
    Next i
    JoinResult = s
End Function

Private Sub CheckL(ByVal Name As String, ByVal Actual As Long, ByVal Expected As Long)
    If Actual = Expected Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected
    End If
End Sub

Private Sub CheckB(ByVal Name As String, ByVal Actual As Boolean, ByVal Expected As Boolean)
    If Actual = Expected Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected
    End If
End Sub

Private Sub CheckS(ByVal Name As String, ByVal Actual As String, ByVal Expected As String)
    If StrComp(Actual, Expected, vbBinaryCompare) = 0 Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got """ & Actual & """, expected """ & Expected & """"
    End If
End Sub

Private Sub CheckD(ByVal Name As String, ByVal Actual As Double, ByVal Expected As Double)
    If Abs(Actual - Expected) < 0.000000001 Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected
    End If
End Sub
