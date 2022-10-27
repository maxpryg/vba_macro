Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Sub swapcolumns()
    ' David McRitchie, 2004-01-05, http://www.mvps.org/dmcritchie/swap.htm
    Dim xlong As Long
    If Selection.Areas.Count <> 2 Then
       MsgBox "Must have exactly two areas for swap." & Chr(10) _
         & "You have " & Selection.Areas.Count & " areas."
       Exit Sub
    End If
    If Selection.Areas(1).Rows.Count <> Cells.Rows.Count Or _
       Selection.Areas(2).Rows.Count <> Cells.Rows.Count Then
       MsgBox "Must select entire Columns, insufficient rows " _
           & Selection.Areas(1).Rows.Count & " vs. " _
           & Selection.Areas(2).Rows.Count & Chr(10) _
           & "You should see both numbers as " & Cells.Rows.Count
       Exit Sub
    End If
    Dim areaSwap1 As Range, areaSwap2 As Range, onepast2 As Range
    '--verify that Area 2 columns follow area 1 columns
    '--so that adjacent single column swap will work.
    If Selection.Areas(1)(1).Column > Selection.Areas(2)(1).Column Then
       Range(Selection.Areas(2).Address & "," & Selection.Areas(1).Address).Select
       Selection.Areas(2).Activate
    End If
    Set areaSwap1 = Selection.Areas(1)
    Set areaSwap2 = Selection.Areas(2)
    Set onepast2 = areaSwap2.Offset(0, areaSwap2.Columns.Count).EntireColumn
    areaSwap2.Cut
    areaSwap1.Resize(, 1).EntireColumn.Insert Shift:=xlShiftToRight
    areaSwap1.Cut
    onepast2.Resize(, 1).EntireColumn.Insert Shift:=xlShiftToRight
    Range(areaSwap1.Address & "," & areaSwap2.Address).Select
    xlong = ActiveSheet.UsedRange.Rows.Count  'correct lastcell
End Sub

Function Substring(Txt, Delimiter, n) As String
Dim x As Variant
    x = Split(Txt, Delimiter)
    If n > 0 And n - 1 <= UBound(x) Then
        Substring = x(n - 1)
    Else
        Substring = ""
    End If
End Function

Sub RemoveLinesWithEmptyCells()
ActiveSheet.AutoFilterMode = False
With Range("A:C")
    .AutoFilter Field:=1, Criteria1:="=Sat.", Operator:=xlOr, Criteria2:="=Sun."
    .AutoFilter Field:=2, Criteria1:="=Sat.", Operator:=xlOr, Criteria2:="=Sun."
    .AutoFilter Field:=3, Criteria1:="="
End With
ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.AutoFilterMode = False
End Sub

Sub FnOpeneWordDoc()
   Dim objWord  As Object
   Dim objDoc   As Object
   
   Dim Path As String
   Dim filename As String
   Dim ttnNubmer As String
   Dim ttnName As String
   
   Path = "D:\TTN\"
   ttnName = Range("C7")
   ttnNumber = Range("C2")
   filename = Path & ttnNumber & "_" & ttnName & ".docx"
     
   Set objWord = CreateObject("Word.Application")
   Set objDoc = objWord.Documents.Open("D:\TTN\TTN_template.docx")
   objWord.Visible = True
   ' objWord.PrintOut (ManualDuplexPrint = True)
   ' objDoc.SaveAs Path & ttnNumber & "_" & ttnName & ".docx"
   objDoc.SaveAs filename

End Sub

Function MySum(vArg1 As Double, vArg2 As Double)
    Dim dblSum As Double
    dblSum = vArg1 + vArg2
    MySum = dblSum
End Function

Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : GoogleTranslate
' DateTime  : 04.09.2013 22:55
' Author    : The_Prist(???????? ???????)
'             http://www.excel-vba.ru
' Purpose   :
'     ??????? ????????? ?????, ????????? ?????? ????????? Google Translate
'     ?????????:
'          sText       - ????? ??? ????????. ??????????????? ????? ??? ?????? ?? ??????.
'          sResLang    - ??? ?????, ?? ??????? ???????????? ???????
'          sSourceLang - ??? ?????, ? ???????? ??????????.
'                        ???? ?? ?????????, Google ????????? ?????????? ???? ?????????? ??????
'                        ???????? 74 ???? ??????.
'---------------------------------------------------------------------------------------
Dim objRegExp As Object
Function GoogleTranslate(ByVal sText As String, ByVal sResLang As String, _
                         Optional ByVal sSourceLang As String = "")

    Dim sGoogleURL As String, sAllTxt As String, sTmpStr As String, sRes As String, sTextToTranslate As String
    Dim lByte As Long, alByteToEncode, li As Long
    Dim asTmp, lPos As Long
    '????????????????? ??? ????????????? ????????? ??????? ??? ????? ????????? ? ?????
    'Application.Volatile True
    If objRegExp Is Nothing Then
        Set objRegExp = CreateObject("VBScript.RegExp")
        With objRegExp
            .MultiLine = True: .ignorecase = True: .Global = True
            .Pattern = "[\n;]"
        End With
    End If
    sTextToTranslate = Application.Trim(objRegExp.Replace(sText, " "))

    With CreateObject("ADODB.Stream")
        .Charset = "utf-8": .Mode = 3: .Type = 2: .Open
        .WriteText sTextToTranslate
        .Flush: .Position = 0: .Type = 1: .Read 3
        alByteToEncode = .Read(): .Close
    End With

    '????????? ????? ? ?????????, ???????? Google
    For li = 0 To UBound(alByteToEncode)
        lByte = alByteToEncode(li)
        Select Case lByte
        Case 32: sTmpStr = "+"
        Case 48 To 57, 65 To 90, 97 To 122: sTmpStr = Chr(alByteToEncode(li))
        Case Else: sTmpStr = "%" & Hex(lByte)
        End Select
        sAllTxt = sAllTxt & sTmpStr
    Next li

    '????????? ?????? ??? Google
    sGoogleURL = "http://translate.google.com.ua/translate_a/t?client=json&text=" & _
                 sAllTxt & "&hl=" & sResLang & "&sl=" & sSourceLang
    sGoogleURL = Replace(sGoogleURL, "\", "/")
    '???????? ????? ?? Google
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", sGoogleURL, "False": .send
        If .statustext = "OK" Then
            sTmpStr = .responsetext
            '???????? ?????? ???????????? ?????
            asTmp = Split(sTmpStr, "{""trans"":""")
            For li = LBound(asTmp) To UBound(asTmp)
                lPos = InStr(1, asTmp(li), """,""orig"":", 1)
                If lPos Then sRes = sRes & Mid(asTmp(li), 1, lPos - 1)
            Next li
            If sRes = "" Then sRes = "?? ??????? ?????????"
        End If
    End With
    GoogleTranslate = sRes
End Function

Option Explicit
 
Sub UpdateWordDoc1()
Dim wdDoc As Object, wdApp As Object
On Error Resume Next
Set wdDoc = CreateObject("D:\TTN\TTN_template.docx")
Set wdApp = wdDoc.Application
wdApp.Visible = True
End Sub

Sub PlusOne()
[C2] = [C2] + 1
End Sub

Sub MinusOne()
[C2] = [C2] - 1
End Sub


Function PropisUkr(n As Double, Optional hryvnias As Variant = False, Optional kopecks As Variant = False) As String

 Nums0 = Array("", "один ", "два ", "три ", "чотири ", "п'ять ", "шість ", "сім ", "вісім ", "дев'ять ")
 Nums1 = Array("", "один ", "два ", "три ", "чотири ", "п'ять ", "шість ", "сім ", "вісім ", "дев'ять ")
 Nums2 = Array("", "десять ", "двадцять ", "тридцять ", "сорок ", "п'ятдесят ", "шістдесят ", "сімдесят ", "вісімдесят ", "дев'яносто ")
 Nums3 = Array("", "сто ", "двісті ", "триста ", "чотириста ", "п'ятсот ", "шістсот ", "сімсот ", "вісімсот ", "дев'ятсот ")
 Nums4 = Array("", "одна ", "дві ", "три ", "чотири ", "п'ять ", "шість ", "сім ", "вісім ", "дев'ять ")
 Nums5 = Array("десять ", "одинадцять ", "дванадцять ", "тринадцять ", "чотирнадцять ", "п'ятнадцять ", "шістнадцять ", "сімнадцять ", "вісімнадцять ", "дев'ятнадцять ")

 Sum = WorksheetFunction.Round(CDbl(n), 2)
 whole = Int(Sum)
 fraq = Format(Round(Abs(Sum - whole) * 100), "00")
 hryvnias = CBool(hryvnias)
 kopecks = CBool(kopecks)

 If Sum < 0 Then
 whole = 0
 fraq = Format(0, "00")
 ed_txt = "Ноль "
 GoTo rrr
 End If

 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 sotmil = Class(n, 9)
 bil = Class(n, 10)

 Select Case bil
 Case 1
 bil_txt = Nums1(bil) & "мільярд "
 Case 2 To 4
 bil_txt = Nums1(bil) & "мільярди "
 Case 5 To 9
 bil_txt = Nums1(bil) & "мільярдів "
 End Select

 Select Case sotmil
 Case 1 To 9
 sotmil_txt = Nums3(sotmil)
 End Select

 Select Case decmil
 Case 1
 mil_txt = Nums5(mil) & "мільйонів "
 GoTo www
 Case 2 To 9
 decmil_txt = Nums2(decmil)
 End Select

 Select Case mil
 Case 0
 If decmil > 0 Then mil_txt = Nums4(mil) & "мільйонів "
 Case 1
 mil_txt = Nums1(mil) & "мільйон "
 Case 2, 3, 4
 mil_txt = Nums1(mil) & "мільйона "
 Case 5 To 9
 mil_txt = Nums1(mil) & "мільйонів "
 End Select

 If decmil = 0 And mil = 0 And sotmil <> 0 Then sotmil_txt = sotmil_txt & "мільйонів "

www:
 sottys_txt = Nums3(sottys)

 Select Case dectys
 Case 1
 tys_txt = Nums5(tys) & "тисяч "
 GoTo eee
 Case 2 To 9
 dectys_txt = Nums2(dectys)
 End Select

 Select Case tys
 Case 0
 If dectys > 0 Then tys_txt = Nums4(tys) & "тисяч "
 Case 1
 tys_txt = Nums4(tys) & "тисячa "
 Case 2, 3, 4
 tys_txt = Nums4(tys) & "тисячі "
 Case 5 To 9
 tys_txt = Nums4(tys) & "тисяч "
 End Select

 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тисяч "

eee:
 sot_txt = Nums3(sot)

 Select Case dec
 Case 1
 ed_txt = Nums5(ed)
 GoTo rrr
 Case 2 To 9
 dec_txt = Nums2(dec)
 End Select

 ed_txt = Nums0(ed)

 If whole < 1 Then ed_txt = "Ноль "

rrr:

 Select Case Class(n, 1) + Class(n, 2) * 10
 Case 1, 21, 31, 41, 51, 61, 71, 81, 91
 grv_text = "гривня"

 Case 2, 3, 4, 22, 23, 24, 32, 33, 34, 42, 43, 44, 52, 53, 54, 62, 63, 64, 72, 73, 74, 82, 83, 84, 92, 93, 94
 grv_text = "гривні"

 Case Else
 grv_text = "гривень"
 End Select

 Select Case fraq
 Case 1, 21, 31, 41, 51, 61, 71, 81, 91
 kop_text = "копiйка"

 Case 2, 3, 4, 22, 23, 24, 32, 33, 34, 42, 43, 44, 52, 53, 54, 62, 63, 64, 72, 73, 74, 82, 83, 84, 92, 93, 94
 kop_text = "копійки"

 Case Else
 kop_text = "копійок"
 End Select

 outstr = bil_txt & sotmil_txt & decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
 If hryvnias Then outstr = outstr & grv_text
 If hryvnias And kopecks Then outstr = outstr & " " & fraq & " " & kop_text

 PropisUkr = UCase(Mid(outstr, 1, 1)) + Mid(outstr, 2)

End Function

Private Function Class(m, i)
 Class = Int(Int(m - (10 ^ i) * Int(m / (10 ^ i))) / 10 ^ (i - 1))
End Function

Private Function ScriptRus(n As Double) As String
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 Nums1 = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
 Nums2 = Array("", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
 Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
 Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
 Nums5 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")

 If n = 0 Then
 ScriptRus = "Ноль"
 Exit Function
 End If
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 sotmil = Class(n, 9)
 mlrd = Class(n, 10)

 If mlrd > 0 Then
 Select Case mlrd
 Case 1
 mlrd_txt = Nums1(mlrd) & "миллиард "
 Case 2, 3, 4
 mlrd_txt = Nums1(mlrd) & "миллиарда "
 Case 5 To 20
 mlrd_txt = Nums1(mlrd) & "миллиардов "
 End Select
 End If
 If (sotmil + decmil + mil) > 0 Then
 sotmil_txt = Nums3(sotmil)

 Select Case decmil
 Case 1
 mil_txt = Nums5(mil) & "миллионов "
 GoTo www
 Case 2 To 9
 decmil_txt = Nums2(decmil)
 End Select

 Select Case mil
 Case 1
 mil_txt = Nums1(mil) & "миллион "
 Case 2, 3, 4
 mil_txt = Nums1(mil) & "миллиона "
 Case 0, 5 To 20
 mil_txt = Nums1(mil) & "миллионов "
 End Select
 End If
www:
 sottys_txt = Nums3(sottys)
 Select Case dectys
 Case 1
 tys_txt = Nums5(tys) & "тысяч "
 GoTo eee
 Case 2 To 9
 dectys_txt = Nums2(dectys)
 End Select

 Select Case tys
 Case 0
 If dectys > 0 Then tys_txt = Nums4(tys) & "тысяч "
 Case 1
 tys_txt = Nums4(tys) & "тысяча "
 Case 2, 3, 4
 tys_txt = Nums4(tys) & "тысячи "
 Case 5 To 9
 tys_txt = Nums4(tys) & "тысяч "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тысяч "
eee:
 sot_txt = Nums3(sot)

 Select Case dec
 Case 1
 ed_txt = Nums5(ed)
 GoTo rrr
 Case 2 To 9
 dec_txt = Nums2(dec)
 End Select

 ed_txt = Nums1(ed)
rrr:

 ScriptRus = mlrd_txt & sotmil_txt & decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
 ScriptRus = UCase(Left(ScriptRus, 1)) & LCase(Mid(ScriptRus, 2, Len(ScriptRus) - 1))
 End Function

Private Function ScriptEng(ByVal Number As Double)
 Dim BigDenom As String, Temp As String
 Dim Count As Integer

 ReDim Place(9) As String
 Place(2) = " Thousand "
 Place(3) = " Million "
 Place(4) = " Billion "
 Place(5) = " Trillion "

 strAmount = Trim(Str(Int(Number)))

 Count = 1
 Do While strAmount <> ""
 Temp = GetHundreds(Right(strAmount, 3))
 If Temp <> "" Then BigDenom = Temp & Place(Count) & BigDenom
 If Len(strAmount) > 3 Then
 strAmount = Left(strAmount, Len(strAmount) - 3)
 Else
 strAmount = ""
 End If
 Count = Count + 1
 Loop
 Select Case BigDenom
 Case ""
 BigDenom = "Zero "
 Case "One"
 BigDenom = "One "
 Case Else
 BigDenom = BigDenom & " "
 End Select
 ScriptEng = BigDenom

End Function

Private Function GetHundreds(ByVal MyNumber)
 Dim result As String
 If Val(MyNumber) = 0 Then Exit Function
 MyNumber = Right("000" & MyNumber, 3)

 If Mid(MyNumber, 1, 1) <> "0" Then
 result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
 End If

 If Mid(MyNumber, 1, 1) <> "0" And (Mid(MyNumber, 2, 1) <> "0" Or Mid(MyNumber, 3, 1) <> "0") Then
 result = result & "And "
 End If

 If Mid(MyNumber, 2, 1) <> "0" Then
 result = result & GetTens(Mid(MyNumber, 2))
 Else
 result = result & GetDigit(Mid(MyNumber, 3))
 End If
 GetHundreds = result
End Function

Private Function GetTens(TensText)
 Dim result As String
 result = ""
 If Val(Left(TensText, 1)) = 1 Then
 Select Case Val(TensText)
 Case 10: result = "Ten"
 Case 11: result = "Eleven"
 Case 12: result = "Twelve"
 Case 13: result = "Thirteen"
 Case 14: result = "Fourteen"
 Case 15: result = "Fifteen"
 Case 16: result = "Sixteen"
 Case 17: result = "Seventeen"
 Case 18: result = "Eighteen"
 Case 19: result = "Nineteen"
 Case Else
 End Select
 Else
 Select Case Val(Left(TensText, 1))
 Case 2: result = "Twenty "
 Case 3: result = "Thirty "
 Case 4: result = "Forty "
 Case 5: result = "Fifty "
 Case 6: result = "Sixty "
 Case 7: result = "Seventy "
 Case 8: result = "Eighty "
 Case 9: result = "Ninety "
 Case Else
 End Select
 result = result & GetDigit _
 (Right(TensText, 1))
 End If
 GetTens = result
End Function
Private Function GetDigit(Digit)
 Select Case Val(Digit)
 Case 1: GetDigit = "One"
 Case 2: GetDigit = "Two"
 Case 3: GetDigit = "Three"
 Case 4: GetDigit = "Four"
 Case 5: GetDigit = "Five"
 Case 6: GetDigit = "Six"
 Case 7: GetDigit = "Seven"
 Case 8: GetDigit = "Eight"
 Case 9: GetDigit = "Nine"
 Case Else: GetDigit = ""
 End Select
End Function
