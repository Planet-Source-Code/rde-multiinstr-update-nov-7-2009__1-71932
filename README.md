<div align="center">

## MultiInStr \(Update Nov 7, 2009\)

<img src="PIC20095251126217921.jpg">
</div>

### Description

This is an improvement on the MultiInStr function that appears in other peoples code now and again ... I don't know who the original author was, so I hope who-ever you are you don't mind ... The original code would search through a string looking for occurences of single characters, while this pair of functions search for single-or-multi character terms within the given string ... Included are MultiInStr and MultiInStrR functions ... Hope someone finds them useful ... Update 25 May - improved versions added thanks to contributions from Kenneth Buckmaster ... Update 7 Nov - Reset string len bug fix in Ken's MultiInStr ... Happy coding
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-multiinstr-update-nov-7-2009__1-71932/archive/master.zip)





### Source Code


<tt>
<p nowrap>
&#160; <br />
<font color="#006600">
&#160;'---------------------------------<br />
&#160; <br />
&#160;' Simple MultiInStr: <br />
&#160;' Always returns 'their' before 'heir' <br />
&#160;' but returns either 'the' or 'their' depending on <br />
&#160;' which term was found first in sTerms array order <br />
</font>
&#160; <br />
<font color="#000099">Function </font><font color="#660000">MultiInStr</font><font color="#330000">(</font><font color="#660000">sSrc </font><font color="#000099">As String</font><font color="#330000">, </font><font color="#660000">sTerms</font><font color="#330000">()</font> <font color="#000099">As String</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">lStart</font> <font color="#000099">As Long</font><font color="#330000"> = 1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">eCompare</font> <font color="#000099">As</font> <font color="#660000">VbCompareMethod</font><font color="#330000"> = </font> <font color="#660000">vbBinaryCompare</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">lRightLimit</font> <font color="#000099">As Long</font><font color="#330000"> = -1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByRef</font> <font color="#660000">lHitItemIndex</font> <font color="#000099">As Long</font><font color="#330000">)</font> </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iPos </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iHit </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iIdx </font><font color="#000099">As Long</font><br />
&#160; <br />
 &#160; <font color="#000099">If </font><font color="#660000">lRightLimit</font> <font color="#330000">= -1</font> <font color="#000099">Then</font> <font color="#660000">lRightLimit</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font><br />
 &#160; <font color="#660000">iHit</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font> <font color="#330000">+ 1</font><br />
&#160; <br />
 &#160; <font color="#000099">For</font> <font color="#660000">iIdx</font> <font color="#330000">=</font> <font color="#000099">LBound(</font><font color="#660000">sTerms</font><font color="#000099">) To UBound(</font><font color="#660000">sTerms)</font><br />
 &#160; &#160; &#160;<font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">lStart</font><font color="#330000">,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
 &#160; &#160; &#160;<font color="#000099">If</font> <font color="#660000">iPos</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#000099">If</font> <font color="#660000">iPos</font> <font color="#330000">&lt;</font> <font color="#660000">iHit</font> <font color="#000099">Then</font> <font color="#660000">iHit</font> <font color="#330000">=</font> <font color="#660000">iPos</font><font color="#330000">:</font> <font color="#660000">lHitItemIndex</font> <font color="#330000">=</font> <font color="#660000">iIdx</font><br />
 &#160; &#160; &#160;<font color="#000099">End If</font><br />
 &#160; <font color="#000099">Next</font><br />
&#160; <br />
 &#160; <font color="#000099">If</font> <font color="#660000">iHit</font> <font color="#330000">&lt;</font> <font color="#660000">Len(sSrc)</font> <font color="#330000">+ 1</font> <font color="#000099">Then</font> <font color="#660000">MultiInStr</font> <font color="#330000">=</font> <font color="#660000">iHit</font><br />
&#160; <br />
<font color="#000099">End Function</font><br />
&#160; <br />
<font color="#006600">
&#160;'---------------------------------<br />
&#160; <br />
&#160;' Comment From: Kenneth Buckmaster<br />
&#160;' It occurred to me that you could avoid searching the<br />
&#160;' whole string length after a term is found<br />
&#160; <br />
&#160;' Also added something you might want in these functions -<br />
&#160;' returns 'the' before 'their' when in the same location<br />
</font>
&#160; <br />
<font color="#000099">Private Declare Sub </font><font color="#660000">CopyMemory </font><font color="#000099">Lib
</font><font color="#000000">"kernel32" </font><font color="#000099">Alias </font><font color="#000000">"RtlMoveMemory" _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#330000">(</font><font color="#660000">pDest
</font><font color="#000099">As Any, </font><font color="#660000">pSrc </font><font color="#000099">As Any, ByVal </font><font color="#660000">lLenB </font><font color="#000099">As Long)</font><br />
&#160; <br />
<font color="#000099">Function </font><font color="#660000">MultiInStr</font><font color="#330000">(</font><font color="#660000">sSrc </font><font color="#000099">As String</font><font color="#330000">, </font><font color="#660000">sTerms</font><font color="#330000">()</font> <font color="#000099">As String</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">lStart</font> <font color="#000099">As Long</font><font color="#330000"> = 1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">eCompare</font> <font color="#000099">As</font> <font color="#660000">VbCompareMethod</font><font color="#330000"> = </font> <font color="#660000">vbBinaryCompare</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByVal</font> <font color="#660000">lRightLimit</font> <font color="#000099">As Long</font><font color="#330000"> = -1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">Optional ByRef</font> <font color="#660000">lHitItemIndex</font> <font color="#000099">As Long</font><font color="#330000">)</font> </font><font color="#000099">As Long</font> <font color="#006600">' Kenneth Buckmaster</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iPos </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iHit </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iIdx </font><font color="#000099">As Long</font><br />
&#160; <br />
 &#160; <font color="#000099">Dim </font><font color="#660000">spointer </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">slenb </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">biggestlen </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">newsearchlen </font><font color="#000099">As Long</font><br />
&#160; <br />
 &#160; <font color="#000099">Dim </font><font color="#660000">bHit </font><font color="#000099">As Boolean</font><br />
&#160; <br />
 &#160; <font color="#660000">slenb </font><font color="#330000">= </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sSrc</font><font color="#330000">)</font><br />
 &#160; <font color="#660000">spointer </font><font color="#330000">= </font><font color="#660000">StrPtr</font><font color="#330000">(</font><font color="#660000">sSrc</font><font color="#330000">) - 4&</font><br />
&#160; <br />
 &#160; <font color="#000099">For </font><font color="#660000">iIdx </font><font color="#330000">=</font> <font color="#000099">LBound(</font><font color="#660000">sTerms</font><font color="#000099">) To UBound(</font><font color="#660000">sTerms)</font><br />
 &#160; &#160; &#160;<font color="#000099">If </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">)) &gt; </font><font color="#660000">biggestlen </font><font color="#000099">Then </font><font color="#660000">biggestlen </font><font color="#330000">= </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">))</font><br />
 &#160; <font color="#000099">Next</font><br />
&#160; <br />
 &#160; <font color="#000099">If </font><font color="#660000">lRightLimit</font> <font color="#330000">= -1</font> <font color="#000099">Then</font> <font color="#660000">lRightLimit</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font><br />
 &#160; <font color="#660000">iHit</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font> <font color="#330000">+ 1</font><br />
&#160; <br />
 &#160; <font color="#000099">For</font> <font color="#660000">iIdx </font><font color="#330000">=</font> <font color="#000099">LBound(</font><font color="#660000">sTerms</font><font color="#000099">) To UBound(</font><font color="#660000">sTerms)</font><br />
 &#160; &#160; &#160;<font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">lStart</font><font color="#330000">,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">If </font><font color="#660000">iPos</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#000099">If </font><font color="#660000">iPos </font><font color="#330000">&lt; </font><font color="#660000">iHit </font><font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#000099">True</font><br />
 &#160; &#160; &#160; &#160; <font color="#000099">ElseIf </font><font color="#660000">iPos </font><font color="#330000">= </font><font color="#660000">iHit </font><font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">)) &lt; </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">lHitItemIndex</font><font color="#330000">))</font><br />
 &#160; &#160; &#160; &#160; <font color="#000099">End If</font><br />
&#160; <br />
 &#160; &#160; &#160; &#160; <font color="#000099">If </font><font color="#660000">bHit</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">iHit </font><font color="#330000">= </font><font color="#660000">iPos</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">lHitItemIndex </font><font color="#330000">= </font><font color="#660000">iIdx</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">newsearchlen </font><font color="#330000">= </font><font color="#660000">iHit </font><font color="#330000">+ <font color="#660000">iHit </font><font color="#330000">+ </font><font color="#660000">biggestlen</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">If </font><font color="#660000">newsearchlen </font><font color="#330000">&lt; </font><font color="#660000">slenb</font> <font color="#000099">Then </font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">CopyMemory </font><font color="#000099">ByVal </font><font color="#660000">spointer</font><font color="#330000">, </font><font color="#660000">newsearchlen</font><font color="#330000">, 4&</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#000099">End If</font><br />
 &#160; &#160; &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#000099">False</font><br />
 &#160; &#160; &#160; &#160; <font color="#000099">End If</font><br />
 &#160; &#160; &#160;<font color="#000099">End If</font><br />
 &#160; <font color="#000099">Next</font><br />
&#160; <br />
 &#160; <font color="#660000">CopyMemory </font><font color="#000099">ByVal </font><font color="#660000">spointer, slenb</font><font color="#330000">, 4&</font><br />
&#160; <br />
 &#160; <font color="#000099">If</font> <font color="#660000">iHit</font> <font color="#330000">&lt;</font> <font color="#660000">Len(sSrc)</font> <font color="#330000">+ 1</font> <font color="#000099">Then</font> <font color="#660000">MultiInStr</font> <font color="#330000">=</font> <font color="#660000">iHit</font><br />
&#160; <br />
<font color="#000099">End Function</font><br />
&#160; <br />
<font color="#006600">
&#160;'---------------------------------<br />
&#160; <br />
&#160;' Simple MultiInStrR:<br />
&#160;' Returns 'heir' before 'their' for reverse search<br />
&#160;' but returns either 'the' or 'their' depending on<br />
&#160;' which term was found first in sTerms array order<br />
</font>
&#160; <br />
<font color="#000099">Function </font><font color="#660000">MultiInStrR</font><font color="#330000">(</font><font color="#660000">sSrc </font><font color="#000099">As String</font><font color="#330000">, </font><font color="#660000">sTerms</font><font color="#330000">()</font> <font color="#000099">As String</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">lRightStart</font> <font color="#000099">As Long</font><font color="#330000"> = -1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">eCompare</font> <font color="#000099">As</font> <font color="#660000">VbCompareMethod</font><font color="#330000"> = </font> <font color="#660000">vbBinaryCompare</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">lLeftLimit</font> <font color="#000099">As Long</font><font color="#330000"> = 1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByRef</font> <font color="#660000">lHitItemIndex</font> <font color="#000099">As Long</font><font color="#330000">)</font> </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iLast </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iPos </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iHit </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iIdx </font><font color="#000099">As Long</font><br />
&#160; <br />
 &#160; <font color="#000099">If </font><font color="#660000">lRightStart</font> <font color="#330000">= -1</font> <font color="#000099">Then</font> <font color="#660000">lRightStart</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font><br />
&#160; <br />
 &#160; <font color="#000099">For</font> <font color="#660000">iIdx</font> <font color="#330000">=</font> <font color="#000099">LBound(</font><font color="#660000">sTerms</font><font color="#000099">) To UBound(</font><font color="#660000">sTerms)</font><br />
 &#160; &#160; &#160;<font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">lLeftLimit</font><font color="#330000">,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">Do Until</font> <font color="#660000">iPos</font> <font color="#330000">= 0</font> <font color="#000099">Or</font> <font color="#660000">iPos</font> <font color="#330000">&gt;</font> <font color="#660000">lRightStart</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iLast</font> <font color="#330000">=</font> <font color="#660000">iPos</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">iLast</font> <font color="#330000">+ 1,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
 &#160; &#160; &#160;<font color="#000099">Loop</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">If</font> <font color="#660000">iLast</font> <font color="#330000">&gt;</font> <font color="#660000">iHit</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iHit</font> <font color="#330000">=</font> <font color="#660000">iLast</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">lHitItemIndex</font> <font color="#330000">=</font> <font color="#660000">iIdx</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">lLeftLimit</font> <font color="#330000">=</font> <font color="#660000">iLast</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iLast</font> <font color="#330000">= 0</font><br />
 &#160; &#160; &#160;<font color="#000099">End If</font><br />
 &#160; <font color="#000099">Next</font><br />
&#160; <br />
 &#160; <font color="#000099">If</font> <font color="#660000">iHit</font> <font color="#000099">Then</font> <font color="#660000">MultiInStrR</font> <font color="#330000">=</font> <font color="#660000">iHit</font><br />
&#160; <br />
<font color="#000099">End Function</font><br />
&#160; <br />
<font color="#006600">
&#160;'---------------------------------<br />
&#160; <br />
&#160;' Comment From: Kenneth Buckmaster<br />
&#160;' Always returns 'heir' before 'their' for reverse search<br />
&#160;' Always returns 'their' before 'the' for reverse search<br />
</font>
&#160; <br />
<font color="#000099">Function </font><font color="#660000">MultiInStrR</font><font color="#330000">(</font><font color="#660000">sSrc </font><font color="#000099">As String</font><font color="#330000">, </font><font color="#660000">sTerms</font><font color="#330000">()</font> <font color="#000099">As String</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">lRightStart</font> <font color="#000099">As Long</font><font color="#330000"> = -1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">eCompare</font> <font color="#000099">As</font> <font color="#660000">VbCompareMethod</font><font color="#330000"> = </font> <font color="#660000">vbBinaryCompare</font><font color="#330000">, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByVal</font> <font color="#660000">lLeftLimit</font> <font color="#000099">As Long</font><font color="#330000"> = 1, _</font><br />
&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#000099">Optional ByRef</font> <font color="#660000">lHitItemIndex</font> <font color="#000099">As Long</font><font color="#330000">)</font> </font><font color="#000099">As Long</font> <font color="#006600">' Kenneth Buckmaster</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iLast </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iPos </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iHit </font><font color="#000099">As Long</font><br />
 &#160; <font color="#000099">Dim </font><font color="#660000">iIdx </font><font color="#000099">As Long</font><br />
&#160; <br />
 &#160; <font color="#000099">Dim </font><font color="#660000">bHit </font><font color="#000099">As Boolean</font><br />
&#160; <br />
 &#160; <font color="#000099">If </font><font color="#660000">lRightStart</font> <font color="#330000">= -1</font> <font color="#000099">Then</font> <font color="#660000">lRightStart</font> <font color="#330000">=</font> <font color="#660000">Len(sSrc)</font><br />
&#160; <br />
 &#160; <font color="#000099">For</font> <font color="#660000">iIdx</font> <font color="#330000">=</font> <font color="#000099">LBound(</font><font color="#660000">sTerms</font><font color="#000099">) To UBound(</font><font color="#660000">sTerms)</font><br />
 &#160; &#160; &#160;<font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">lLeftLimit</font><font color="#330000">,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">Do Until</font> <font color="#660000">iPos</font> <font color="#330000">= 0</font> <font color="#000099">Or</font> <font color="#660000">iPos</font> <font color="#330000">&gt;</font> <font color="#660000">lRightStart</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iLast</font> <font color="#330000">=</font> <font color="#660000">iPos</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iPos</font> <font color="#000099">= InStr(</font><font color="#660000">iLast</font> <font color="#330000">+ 1,</font> <font color="#660000">sSrc</font><font color="#330000">,</font> <font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">), <font color="#660000">eCompare</font><font color="#330000">)</font><br />
 &#160; &#160; &#160;<font color="#000099">Loop</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">If</font> <font color="#660000">iLast</font> <font color="#330000">&gt;</font> <font color="#660000">iHit</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#000099">True</font><br />
 &#160; &#160; &#160;<font color="#000099">ElseIf </font><font color="#660000">iLast </font><font color="#330000">= </font><font color="#660000">iHit </font><font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">iIdx</font><font color="#330000">)) &gt; </font><font color="#660000">LenB</font><font color="#330000">(</font><font color="#660000">sTerms</font><font color="#330000">(</font><font color="#660000">lHitItemIndex</font><font color="#330000">))</font><br />
 &#160; &#160; &#160;<font color="#000099">End If</font><br />
&#160; <br />
 &#160; &#160; &#160;<font color="#000099">If </font><font color="#660000">bHit</font> <font color="#000099">Then</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iHit</font> <font color="#330000">=</font> <font color="#660000">iLast</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">lHitItemIndex</font> <font color="#330000">=</font> <font color="#660000">iIdx</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">lLeftLimit</font> <font color="#330000">=</font> <font color="#660000">iLast</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">iLast</font> <font color="#330000">= 0</font><br />
 &#160; &#160; &#160; &#160; <font color="#660000">bHit </font><font color="#330000">= </font><font color="#000099">False</font><br />
 &#160; &#160; &#160;<font color="#000099">End If</font><br />
 &#160; <font color="#000099">Next</font><br />
&#160; <br />
 &#160; <font color="#000099">If</font> <font color="#660000">iHit</font> <font color="#000099">Then</font> <font color="#660000">MultiInStrR</font> <font color="#330000">=</font> <font color="#660000">iHit</font><br />
&#160; <br />
<font color="#000099">End Function</font><br />
&#160; <br />
&#160; <br />
</p>
</tt>

