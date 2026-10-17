# EntityResolutionVBA
Groups records that refer to the same person (entity resolution) in VBA, using the name and, optionally, date of birth and an ID. It works in Excel, Access, Word and other Office applications.

```vba
Dim ids As Variant
ids = ER_Resolve(Array("John Smith", "John Q Smith", "John Quincy Smith", _
                       "William Jones", "Bill Jones", "Mary Smith"))
' ids = (1, 1, 1, 2, 2, 3)
```

## Installation
The matching builds on two other libraries. In the VBA editor (Alt+F11) choose **File > Import File...** and import:

| File | From |
|---|---|
| `EntityResolution.bas` | this repository |
| `JaroWinkler.bas`, `DamerauLevenshtein.bas`, `Phonetic.bas` | [TextDistanceVBA](https://github.com/Calmaroon/TextDistanceVBA) |
| `Phonetic_Encodings/Encode_DoubleMetaphone.bas` | [PhoneticVBA](https://github.com/Calmaroon/PhoneticVBA) |

Optionally import `tests/EntityResolutionTests.bas` and run `Test_EntityResolution` to check the setup; results appear in the Immediate window (Ctrl+G).

## Functions

| Function | Result |
|---|---|
| `ER_Resolve(Names, [DOBs], [IDs], [Threshold], [Blocking], [ReturnNames])` | Entity number for each name (same shape as `Names`; blanks give `""`). `ReturnNames:=True` returns the fullest name in each group instead, e.g. `"John Quincy Smith"` for all three Johns |
| `ER_MatchScore(Name1, Name2, [DOB1], [DOB2], [ID1], [ID2])` | 0 to 1; 0 means the two can't be the same person |
| `ER_IsMatch(..., [Threshold])` | `ER_MatchScore >= Threshold` (default 0.85) |
| `ER_Explain(...)` | Score with the reason for each part, e.g. `0.975: first nickname (BILL~WILLIAM); middle -; last exact (JONES)` |
| `ER_ParseName(Name, [Part])` | Cleaned name, or `"FIRST"`, `"MIDDLE"`, `"LAST"`, `"SUFFIX"` |
| `ER_AddNicknames "Ivan\|Vanya"` / `ER_ResetNicknames` | Extend the built-in nickname list (about 175 groups of common English names) or restore it |

## Using it with a table
Give `ER_Resolve` a column of names and, optionally, matching columns of dates of birth and IDs, in the same row order. Two people with the same name but different dates of birth are kept apart:

| Name | DOB | Entity |
|---|---|---|
| Robert Brown | 1980-05-12 | 1 |
| Bob Brown | 1980-05-12 | 1 |
| Robert Brown | 1990-02-03 | 2 |
| Brown, Robert J | 1990-02-03 | 2 |

### Excel
With names in column A and dates of birth in column B, type this in C2 (Excel 365 fills it down):

```
=ER_Resolve(A2:A500, B2:B500)
```

- First and last names in separate columns (A, B), DOB in C: `=ER_Resolve(A2:A500 & " " & B2:B500, C2:C500)`
- An ID column as well (D): `=ER_Resolve(A2:A500, B2:B500, D2:D500)`
- The fullest name in each group instead of a number: `=ER_Resolve(A2:A500, B2:B500, , , , TRUE)`

For older Excel, or to write plain values instead of a live formula:

```vba
Sub ResolvePeople()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("C2:C" & lastRow).Value = _
        ER_Resolve(ws.Range("A2:A" & lastRow), ws.Range("B2:B" & lastRow))
End Sub
```

### Access
`ER_Resolve` has to see all the rows at once, so it can't be called row by row in a query. Read the table, resolve it and write the numbers back to a number field (here `EntityID`):

```vba
Sub ResolvePeople()
    Dim rs As DAO.Recordset, n As Long, i As Long
    Dim names() As Variant, dobs() As Variant, ids As Variant

    Set rs = CurrentDb.OpenRecordset("SELECT FullName, DOB, EntityID FROM People", dbOpenDynaset)
    If rs.EOF Then Exit Sub
    rs.MoveLast: n = rs.RecordCount: rs.MoveFirst
    ReDim names(1 To n): ReDim dobs(1 To n)
    For i = 1 To n
        names(i) = rs!FullName          ' or: rs!FirstName & " " & rs!LastName
        dobs(i) = rs!DOB                ' Null is fine
        rs.MoveNext
    Next i

    ids = ER_Resolve(names, dobs)

    rs.MoveFirst
    For i = 1 To n
        rs.Edit
        If ids(i) = "" Then rs!EntityID = Null Else rs!EntityID = ids(i)
        rs.Update
        rs.MoveNext
    Next i
    rs.Close
End Sub
```

`ER_MatchScore`, `ER_IsMatch`, `ER_Explain` and `ER_ParseName` work on one pair at a time, so they can be used directly in queries, e.g. `SELECT ER_MatchScore([Name1], [Name2]) FROM ...`.

## How records are compared

1. **Parsing.** Each name is split into first, middle, last and suffix. Parsing handles "Last, First" order, titles (Dr, Mrs), credentials (PhD), accents, punctuation, and surname particles (`van Beethoven` becomes `VANBEETHOVEN`).
2. **Parts.** Each part is scored as one of:
   - exact (1)
   - nickname, e.g. Bill = William (0.95)
   - one-letter typo by Damerau-Levenshtein, or two letters for names of 8+ characters (0.92 for surnames, 0.9 for given names)
   - sounds alike: same Double Metaphone or MRA code, or close Editex, with Jaro-Winkler >= 0.8 (0.9 / 0.88)
   - an initial, e.g. Q = Quincy (0.85)

   A part that clearly disagrees is a **conflict**: John vs Mary, Eric vs Erica, middle Q vs R, Jr vs Sr, a different date of birth, a different ID.
3. **Score.** The score is the average of the first- and last-name scores, plus +0.03 for a matching middle name (+0.01 for a matching initial) and +0.05 for the same date of birth, capped at 1. A missing first or last name counts 0.6. A conflict gives 0. Equal IDs give 1 whatever the names.
4. **Grouping.** Pairs scoring at least `Threshold` are linked, strongest first. A link is skipped if any record in one group conflicts with any record in the other, so `John Q Smith` and `John R Smith` stay apart even though both match `John Smith`.

Dates of birth can be dates, `"YYYY-MM-DD"` text or `YYYYMMDD` numbers; other date text is read with your system's date format. A swapped day/month or a single wrong digit doesn't count as a conflict, but it gets no bonus either. A record with no date of birth is matched on its name alone, so if two people share that name it joins whichever group is linked first; review such rows. IDs (SSN, email, employee number…) are compared ignoring case and spaces.

With `Blocking` on (the default), only records whose surnames sound alike, share their first 3 letters or letters 2–4 (catching a wrong first letter, as in Chang / Zhang), start with the same sounds, or share an ID are compared. This keeps large lists fast. Records that share a common surname are still compared with each other, so time grows with the square of the largest such group. Turn `Blocking` off to compare every pair.

### Limitations
- A surname change (e.g. after marriage) only matches when an ID links the records.
- First and last name swapped without a comma ("Smith John") is not recognised.
- The nickname list covers common English names; add others with `ER_AddNicknames`.

## Tests
`python3 tools/run_tests.py` runs the tests and a few worksheet formulas in headless LibreOffice (VBA compatibility mode). It needs LibreOffice with its Python bridge (`python3-uno`) and checkouts of TextDistanceVBA and PhoneticVBA, either next to this repository or pointed to by `TEXTDISTANCEVBA_DIR` and `PHONETICVBA_DIR`.
