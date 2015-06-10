<%
' ASP Profiler v2.5
' Copyright © 2001-2015 Zafer Barutcuoglu. All Rights Reserved.
'
' ASP Profiler is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
' 
' ASP Profiler is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with ASP Profiler; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
' 
' Visit http://aspprofiler.sourceforge.net/ for more information.
'
' - 20010507 Created.
' - 20120424 CInt -> CLng (thanks to Diane Connors).
' - 20130813 Fixed header bug, added newline conversion (thanks to Tomas Wallentinus).
' - 20150209 Added meta tag for IE10+ (thanks to Peter Chr. Gram).

Option Explicit

''Optional check for my staging server. You will probably never use this.
'CheckServer
'Sub CheckServer()
'   Dim objHTTP, s
'   Set objHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'   objHTTP.Open "GET", "http://" & Request.ServerVariables("SERVER_NAME") & "/serverknt.asp", False
'   objHTTP.Send
'   s = objHTTP.ResponseText
'   Set objHTTP = Nothing
'   If InStr(1, s, "stage") = 0 Then Response.Redirect "/"
'End Sub

Const sInternalCount = "#LINECOUNT#"
Const sInternalFileName = "#FILENAME#"

Sub FindInclude(sFile, iPos, iLength, sIncPath)
   Dim r, ms, m, i, j, s
   Set r = New RegExp
   r.Pattern = "<!--\s*#include\s+(file|virtual)\s*=\s*"".+""\s*-->"
   r.IgnoreCase = True
   r.Global = False
   Set ms = r.Execute(sFile)
   If ms.Count = 0 Then 
      iPos = 0
   Else
      For Each m In ms
         iPos = m.FirstIndex + 1
         iLength = m.Length
         s = m.Value
         i = Instr(1, s, """")
         j = Instr(i+1, s, """")
         sIncPath = Mid(s, i+1, j-i-1)
         i = Instr(1, s, "virtual", vbTextCompare)
         If i > 0 And Left(sIncPath,1) <> "/" Then sIncPath = "/" & sIncPath
      Next
   End If
End Sub

Function ResolveInclude(sIncPath, dictLimits)
   Dim sIncText, f
   On Error Resume Next
   Set f = fso.OpenTextFile(Server.MapPath(sIncPath))
   If Err Then Response.Write "Cannot open " & sIncPath : Response.End
   sIncText = ConvertCrLf(f.ReadAll)
   ResolveAllIncludes sIncText, sIncPath, dictLimits
   ResolveInclude = sIncText
End Function

Sub ResolveAllIncludes(sFile, sWebPath, dictLimits)
   Dim iPos, iLength, sIncText, sIncPath
   Dim sBefore, iIncFirstLine, iTotalIncLines, iIncLines, dictIncLimits, arrDictItem
   
   iTotalIncLines = 0
   dictLimits.Add dictLimits.Count, Array(1, sWebPath, 1)
   
   FindInclude sFile, iPos, iLength, sIncPath
   Do Until iPos = 0
      If Left(sIncPath,1) <> "/" Then sIncPath = fso.GetParentFolderName(sWebPath) & "/" & sIncPath
      sBefore = Left(sFile, iPos-1)
      
      iIncFirstLine = CountLines(sBefore)
      Set dictIncLimits = Server.CreateObject("Scripting.Dictionary")
      
      sIncText = ResolveInclude(sIncPath, dictIncLimits)
      sFile = sBefore & sIncText & Mid(sFile, iPos + iLength)
      
      For Each arrDictItem In dictIncLimits.Items
         arrDictItem(0) = arrDictItem(0) + iIncFirstLine - 1
         dictLimits.Add dictLimits.Count, arrDictItem
      Next
      iIncLines = CountLines(sIncText)
      dictLimits.Add dictLimits.Count, Array(iIncFirstLine + iIncLines, sWebPath, iIncFirstLine - iTotalIncLines + 1)
      iTotalIncLines = iTotalIncLines + iIncLines - 1
      
      FindInclude sFile, iPos, iLength, sIncPath
   Loop
End Sub

Function LoadFile(sFilePath)
   Dim ts
   Set ts = fso.OpenTextFile(sFilePath)
   LoadFile = ConvertCrLf(ts.ReadAll)
   ts.Close
End Function

Sub SaveFile(sFilePath, sData)
   Dim ts
   Set ts = fso.OpenTextFile(sFilePath, 2, True) 'ForWriting, Create
   ts.Write sData
   ts.Close
End Sub

Function GetProfileURL(ByVal sURL)
   GetProfileURL = fso.GetParentFolderName(sURL) & "/" &  fso.GetBaseName(sURL) & ".profile." & fso.GetExtensionName(sURL)
End Function

Function ConvertCrLf(ByVal sText)
   sText = Replace(sText, vbCrLf, vbCr)
   sText = Replace(sText, vbLf, vbCr)
   sText = Replace(sText, vbCr, vbCrLf)
   ConvertCrLf = sText
End Function

Function ProfileCode(ByVal sCode, ByVal sBaseName)
   Dim i, j, bFirst, iLines, iFirstLine
   Dim sOut, iPreviousLine
   
   'Get very base name of sBaseName
   i = InStr(1, sBaseName, ".")
   If i > 0 Then sBaseName = Left(sBaseName, i - 1)

   iLines = CountLines(sCode)
   sOut = ""
   bFirst = True
   iPreviousLine = -1
   j = 1
   i = InStr(1, sCode, "<" & "%")

   Do While i > 0
      sOut = sOut & Mid(sCode, j, i - j + 2)
      j = InStr(i + 2, sCode, "%" & ">")
      If j = 0 Then
         Response.Write "Unclosed ASP tag!"
         Exit Do
      End If
      
      iFirstLine = CountLines(Left(sCode, i))
      sOut = sOut & ProfileCodeBlock(Mid(sCode, i + 2, j - i - 2), iFirstLine, iPreviousLine, bFirst)
      i = InStr(j, sCode, "<" & "%")
   Loop
   sOut = sOut & Mid(sCode, j)
   
   If bFirst Then sOut = sOut & "<" & "%" & GetProfileHeader & "%" & ">" & vbCrLf

   AddReportingFooter sOut

   sOut = Replace(sOut, "Response.End", "Profiler_End", 1, -1, 1)
   sOut = Replace(sOut, "Response.Redirect", "Profiler_End '", 1, -1, 1)
   sOut = Replace(sOut, "Server.Transfer", "Profiler_End '", 1, -1, 1)
   sOut = Replace(sOut, "*Profiler_End*", "Response.End", 1, -1, 1)
   sOut = Replace(sOut, sInternalCount, iLines, 1, -1, 1)
   'sOut = Replace(sOut, sInternalFileName, sBaseName)
   sOut = Replace(sOut, sInternalFileName, "", 1, -1, 1)
   ProfileCode = sOut
End Function

Function GetProfileHeader()
   Dim sCode
   sCode = "' Intermediate file created by ASP Profiler at " & Now & ". Delete after use." & vbCrLf
   sCode = sCode & "Dim Tpr_" & sInternalFileName & "_N(" & sInternalCount & "), Tpr_" & sInternalFileName & "_T(" & sInternalCount & "), Tpr_" & sInternalFileName & "_T1" & vbCrLf
   sCode = sCode & "Sub Profiler_End()" & vbCrLf
   sCode = sCode & "Response.Clear" & vbCrLf
   'sCode = sCode & "Response.Write ""Line:Int|Count:Int|Time:Int|Percent|File|FileLine|Code"" & vbCrLf" & vbCrLf
   sCode = sCode & "For Tpr_" & sInternalFileName & "_T1 = 1 To UBound(Tpr_" & sInternalFileName & "_N)" & vbCrLf
   sCode = sCode & " Response.Write Tpr_" & sInternalFileName & "_T1 & ""|0"" & Tpr_" & sInternalFileName & "_N(Tpr_" & sInternalFileName & "_T1) & ""|0"" & Tpr_" & sInternalFileName & "_T(Tpr_" & sInternalFileName & "_T1) & vbCrLf" & vbCrLf
   sCode = sCode & "Next" & vbCrLf
   sCode = sCode & "*Response.End*" & vbCrLf
   sCode = sCode & "End Sub" & vbCrLf
   GetProfileHeader = sCode
End Function

Function ProfileCodeBlock(ByVal sBlock, ByVal iFirstLine, ByRef iPreviousLine, ByRef bFirst)
   Dim arrLines, i, j, sLine, iLine
   Dim bSkipLine, bSkipTime, bSkipAbsolute
   Dim bUnfinishedLine, iBeginningLine 'for "_" lines
   arrLines = Split(sBlock, vbCrLf)
   sBlock = ""
   iBeginningLine = -1
   For i = 0 To UBound(arrLines)
      sLine = Trim(Replace(arrLines(i), vbTab, ""))
      If Len(sLine) > 0 Then
         CheckLine sLine, bSkipLine, bSkipTime, bSkipAbsolute, bUnfinishedLine
         If Not bSkipAbsolute And bFirst Then
            sBlock = sBlock & GetProfileHeader
            bFirst = False
         End If
         If Not bSkipLine Then
            iLine = iFirstLine + i
            
            If bUnfinishedLine Then
               'in a "_" sequence
               If iBeginningLine < 0 Then iBeginningLine = iLine 'first
               sBlock = sBlock & ProfileCodeLine(arrLines(i), iBeginningLine, (iBeginningLine <> iLine) Or (iLine = iPreviousLine), (iBeginningLine <> iLine) Or bSkipTime, True)
            Else
               If iBeginningLine > 0 Then
                  'last line of a "_" sequence
                  sBlock = sBlock & ProfileCodeLine(arrLines(i), iBeginningLine, True, True, bSkipTime)
                  iBeginningLine = -1
               Else
                  'profile
                  sBlock = sBlock & ProfileCodeLine(arrLines(i), iLine, (iLine = iPreviousLine), bSkipTime, bSkipTime)
               End If
            End If
         Else
            'ignore
            sBlock = sBlock & arrLines(i)
         End If
      Else
         'ignore
         sBlock = sBlock & arrLines(i)
      End If
      If i < UBound(arrLines) Then sBlock = sBlock & vbCrLf
   Next
   ProfileCodeBlock = sBlock
   iPreviousLine = iLine
End Function

Function GetBeginProfileLine(ByVal iLine, ByVal bSkipCount, ByVal bSkipStartTime)
   Dim sOut
   sOut = ""
   If Not bSkipCount Then sOut = sOut & "Tpr_" & sInternalFileName & "_N(" & iLine & ") = Tpr_" & sInternalFileName & "_N(" & iLine & ") + 1: "
   If Not bSkipStartTime Then sOut = sOut & "Tpr_" & sInternalFileName & "_T1=Timer" & vbCrLf
   
   GetBeginProfileLine = sOut
End Function

Function GetProfileLine(ByVal sLine)
   If Left(Trim(Replace(sLine, vbTab, "")), 1) = "=" Then
      GetProfileLine =  "%" & "><" & "%" & sLine & "%" & "><" & "%" & vbCrLf
   Else
      GetProfileLine = sLine & vbCrLf
   End If
End Function

Function GetEndProfileLine(ByVal iLine, ByVal bSkipEndTime)
   Dim sOut
   sOut = ""
   If Not bSkipEndTime Then sOut = sOut & "Tpr_" & sInternalFileName & "_T(" & iLine & ") = Tpr_" & sInternalFileName & "_T(" & iLine & ") + 256 * (Timer - Tpr_" & sInternalFileName & "_T1)" & vbCrLf
   
   GetEndProfileLine = sOut
End Function

Function ProfileCodeLine(ByVal sLine, ByVal iLine, ByVal bSkipCount, ByVal bSkipStartTime, ByVal bSkipEndTime)
   Dim sOut
   
   sOut = GetBeginProfileLine(iLine, bSkipCount, bSkipStartTime)
   sOut = sOut & GetProfileLine(sLine)
   sOut = sOut & GetEndProfileLine(iLine, bSkipEndTime)
   
   ProfileCodeLine = sOut
End Function

Function CountLines(sCode)
   CountLines = UBound(Split(sCode, vbCrLf)) + 1
End Function

Sub AddReportingFooter(ByRef sCode)
   
   sCode = sCode & "<" & "%Profiler_End%" & ">"
End Sub

Sub CheckLine(ByVal sLine, ByRef bSkipLine, ByRef bSkipTime, ByRef bSkipAbsolute, ByRef bUnfinishedLine)
   Dim i, sFirst
   
   bSkipAbsolute = False 'to skip Option Explicit etc before writing Dim.
   bUnfinishedLine = (Right(RTrim(sLine),1) = "_")
   
   '@ directives
   If Left(sLine, 1) = "@" Then '@language = vbscript etc.
      bSkipLine = True
      bSkipAbsolute = True
      Exit Sub
   End If
   
   'comments
   If Left(sLine, 1) = "'" Then 
      bSkipLine = True
      Exit Sub
   End If
   
   'block keywords
   i = InStr(1, sLine, " ")
   If i = 0 Then sFirst = sLine Else sFirst = Left(sLine, i - 1)
   sFirst = LCase(sFirst)
   Select Case sFirst
      Case "option"
         bSkipAbsolute = True
         bSkipLine = True
         Exit Sub
      Case "case", "dim", "public", "private", "const", "class", "sub", "function", "property", "static", "end"
         bSkipLine = True
         Exit Sub
      Case "select", "class"
         bSkipLine = False
         bSkipTime = True
         Exit Sub
   End Select
   
   bSkipLine = False
   bSkipTime = False
End Sub

Sub WriteDefaultPage()
  WriteHeader
%>
   <form method="post" name="frmProfile">
      <input type="hidden" name="cmd" value="">
      <b>URL:</b> http://<%=Request.ServerVariables("SERVER_NAME")%>/ <input type="text" name="path" value="index.asp"><br>
      <br>
      <input type="button" value="Create Intermediate File" onclick="this.form.cmd.value='create';this.form.submit();"> if your server has Write access from within ASP files.<br>
      <input type="button" value="View Intermediate Source" onclick="this.form.cmd.value='view';this.form.submit();"> if that fails, to save and upload the file yourself.<br>
   </form>
<%
  WriteFooter
End Sub

Sub WriteHeader()
%>
  <basefont face="Arial">
  <font size=+3><b>ASP Profiler v2.4</b></font><br>
<%
End Sub

Sub WriteFooter()
%>
  <font size=-1>
  Copyright © 2001-2015 Zafer Barutcuoglu. All Rights Reserved.<br>
  Visit <a href="http://aspprofiler.sourceforge.net/">http://aspprofiler.sourceforge.net/</a> for more information.<br>
  </font>
<%
End Sub

'
' Main code
'

Dim fso, sWebPath, sFilePath, sCmd
Set fso = Server.CreateObject("Scripting.FileSystemObject")

sWebPath = Request("path")
sCmd = Request("cmd")

'Default page
If Len(sWebPath) = 0 Or Len(sCmd) = 0 Then
   WriteDefaultPage
   Response.End
End If

sWebPath = Replace(sWebPath, "\", "/")
If InStr(1, sWebPath, "?") > 0 Then sWebPath = Left(sWebPath, InStr(1, sWebPath, "?") - 1)
If Left(sWebPath, 1) <> "/" Then sWebPath = "/" & sWebPath

sFilePath = Server.MapPath(sWebPath)
If Not fso.FileExists(sFilePath) Then
   Response.Write "<font color='red'>ERROR: Path not found: <b>" & sWebPath & "</b></font><br>"
   WriteDefaultPage
   Response.End
End If

Dim sData, sNewData, dictLimits
sData = LoadFile(sFilePath)

Set dictLimits = Server.CreateObject("Scripting.Dictionary")
ResolveAllIncludes sData, sWebPath, dictLimits

sNewData = ProfileCode(sData, fso.GetBaseName(sFilePath))

Dim sProfileURL, sProfilePath
sProfileURL = GetProfileURL(sWebPath)
sProfilePath = Server.MapPath(sProfileURL)

If sCmd = "create" Then
  SaveFile sProfilePath, sNewData
End If

%>

<html>
<meta http-equiv="x-ua-compatible" content="IE=10">
<head>
<title>ASP Profiler v2.5</title>
<!-- ZTable BEGIN -->
<SCRIPT LANGUAGE="VBSCRIPT">
Option Explicit

Class ZTable

Dim bSortAsc, iSortFld, LastURL

Public Items, RowDelim, FieldDelim, ElementID, Header, Footer, XMLWrap, RowFormat

Private Sub Class_Initialize
   RowDelim = vbCrLf
   FieldDelim = ","
   Header = "<table>"
   Footer = "</table>"
   RowFormat = ""
   XMLWrap = False
End Sub

Public Sub Refresh()
   Window.Status = "Displaying..."
   Dim s, i, j, arrLines, r
   ReDim arrLines(UBound(Items))
   s = Header

   If Len(RowFormat) > 0 Then
      For i=0 to UBound(Items)
         r = RowFormat
         For j=0 To UBound(Items(i))
            r = Replace(r, "%" & j & "%" , Items(i)(j))
         Next
         arrLines(i) = r
      Next
      s = s & Join(arrLines, "")
   Else
      For i=0 to UBound(Items)
         arrLines(i) = Join(Items(i),"</td><td>")
      Next
      s = s & "<tr><td>" & Join(arrLines, "</td></tr><tr><td>") & "</td></tr>"
   End If

   s = s & Footer
   Execute ElementID & ".innerHTML = s"
   Window.Status = "Done"
End Sub

Public Sub Sort(iField, bAscending)
   Window.Status = "Sorting..."
   bSortAsc = CBool(bAscending)
   iSortFld = iField
   QuickSort 0, UBound(Items)
   Refresh
End Sub

Private Sub InsertionSort(iFirst, iLast)
   Dim i, j, v
   For i=iFirst+1 To iLast
      v = Items(i)
      j = i
      Do While (bSortAsc Eqv (v(iSortFld) < Items(j-1)(iSortFld)))
         Items(j) = Items(j-1)
         j = j - 1
         If j = iFirst Then Exit Do
      Loop
      Items(j) = v
   Next
End Sub

Private Sub QuickSort(iFirst, iLast)
   Dim pivot, i, j, k
   If iLast - iFirst < 7 Then
      InsertionSort iFirst, iLast
   Else
      k = Int((iFirst + iLast) / 2)
      If (Items(iFirst)(iSortFld) > Items(k)(iSortFld)) Eqv bSortAsc Then SwapItems iFirst, k
      If (Items(iFirst)(iSortFld) > Items(iLast)(iSortFld)) Eqv bSortAsc Then SwapItems iFirst, iLast
      If (Items(k)(iSortFld) > Items(iLast)(iSortFld)) Eqv bSortAsc Then SwapItems k, iLast
      SwapItems k, iLast-1
      pivot = Items(iLast-1)(iSortFld)
      i = iFirst
      j = iLast-1
      Do
         Do: i = i + 1: Loop While (Items(i)(iSortFld) <= pivot) Eqv bSortAsc
         Do: j = j - 1: Loop While (Items(j)(iSortFld) >= pivot) Eqv bSortAsc
         If i < j Then
            SwapItems i, j
         Else
            Exit Do
         End If
      Loop
      SwapItems i, iLast-1
      
      QuickSort iFirst, i-1
      QuickSort i+1, iLast
   End If
End Sub

Private Sub SwapItems(i, j)
   Dim tmp
   tmp = Items(j)
   Items(j) = Items(i)
   Items(i) = tmp
End Sub

Public Sub Reload()
   If Len(LastURL) > 0 Then LoadURL LastURL
End Sub

Public Sub LoadURL(url)
   LastURL = url
   Window.Status = "Loading..."
   Dim s
   If XMLWrap Then
      Dim objXML
      Set objXML = CreateObject("Microsoft.XMLDOM")
      objXML.Async = False
      objXML.Load url
      s = objXML.documentElement.Text
      Set objXML = Nothing
   Else
      Dim objHTTP
      Set objHTTP = CreateObject("Microsoft.XMLHTTP")
      objHTTP.Open "GET", url, False
      objHTTP.Send
      s = objHTTP.responseText
      Set objHTTP = Nothing
   End If
   LoadData s
End Sub

Public Sub LoadData(s)
   ParseFile s
   Refresh
End Sub

Private Sub ParseFile(s)
   On Error Resume Next
   Window.Status = "Parsing..."
   Do While Left(s, Len(RowDelim)) = RowDelim
      s = Mid(s, Len(RowDelim) + 1)
   Loop
   Do While Left(s, 2) = "<?"
      s = Mid(s, InStr(1, s, "?>")+2)
   Loop
   Do While Right(s, Len(RowDelim)) = RowDelim
      s = Left(s, Len(s) - Len(RowDelim))
   Loop
   Dim i, j
   Items = Split(s, RowDelim)
   For i=0 To UBound(Items)
      Items(i) = Split(Items(i), FieldDelim)
      For j=0 to UBound(Items(i))
         Items(i)(j) = Eval(Items(i)(j))
      Next
   Next
End Sub

End Class
-->
</SCRIPT>
<!-- ZTable END -->

<SCRIPT LANGUAGE=VBSCRIPT>
<!--

Dim s, k, z, lastSort
function initSrc()
<%
   Dim arrData, i, t
   arrData = Split(sData, vbCrLf)
%>
   ReDim s(<%=UBound(arrData)+1%>)
<%
   For i = 0 To UBound(arrData)
      t = Replace(arrData(i),"""","""""")
      't = Replace(t, "-->", "--"" & "">")
      Response.Write "s(" & i+1 & ")=""" & Replace(Server.HTMLEncode(t), " ", "&nbsp;") & """" & vbCrLf
   Next
%>
   ReDim k(<%=dictLimits.Count-1%>)
<%
   i = 0
   For Each t In dictLimits.Items
      Response.Write "k(" & i & ")= Array(" & t(0) & ", """ & t(1) & """, " & t(2) & ")" & vbCrLf
      i = i + 1
   Next
%>

   lastSort = 2

   Set z = New ZTable
   z.ElementID = "divResult"
   z.FieldDelim = "|"
   z.Header = "<table border=1><thead><tr><td onclick='SortTable(0)'><b><u><font color=blue>Line#</font></u></b></td><td onclick='SortTable(1)'><b><u><font color=blue>Count</font></u></b></td><td onclick='SortTable(2)'><b><u><font color=blue>Time</font></u></b></td><td><b>Time%</b></td><td><b>File</b></td><td><b>FLine#</b></td><td><b>Code</b></td></tr></thead><tbody>"
   z.Footer = "</tbody></table>"
   z.RowFormat = "<tr bgColor=%7%><td>%0%</td><td>%1%</td><td>%2%</td><td>%3%</td><td style='font-size:7pt'>%4%</td><td>%5%</td><td style='font-size:7pt'>%6%</td></tr>"
end function

sub SortTable(iField)
   lastSort = iField
   Select Case iField
      Case 0
         z.Sort iField,1 'asc
      Case 1, 2
         z.Sort iField,0 'desc
   End Select
end sub

sub FindFilePos(j, byref sFile, byref iFileLine)
   Dim i, arr
   arr = k(0)
   For i = 1 to <%=dictLimits.Count-1%>
      arr = k(i)
      If arr(0) > CLng(j) Then arr = k(i-1): Exit For
   Next
   sFile = arr(1)
   iFileLine = arr(2) + j - arr(0)
end sub

function completeTable()
   Dim r, iTotalTime, i, j, sFile, iFileLine
   iTotalTime = 0
   i = 0
   For i=0 To UBound(z.Items)
      r = z.Items(i)
      ReDim Preserve r(7)
      j = r(0)
      r(6) = s(j) 'code
      iTotalTime = iTotalTime + r(2)
      FindFilePos j, sFile, iFileLine
      r(4) = sFile
      r(5) = iFileLine
      If r(1) > 0 Then r(7) = "#DDDDDD" 'Else r(7) = "" 'bgColor
      z.Items(i) = r
   Next
   If iTotalTime > 0 Then 
      divSummary.innerText = "Profile Time: " & iTotalTime & " cycles (" & CLng(iTotalTime * 1000 / 256) & " ms)"
      For i=0 To UBound(z.Items)
         r = z.Items(i)
         r(3) = Round(z.Items(i)(2) * 100 / iTotalTime)
         z.Items(i) = r
      Next
   End If
end function

function cmdRun_onclick()
   z.LoadURL hiddenProfileURL.Value & "?" & txtDataURL.Value
   completeTable
   SortTable lastSort
end function


function cmdDownloadTime_onclick()
   Dim objHTTP, t, size
   Set objHTTP = CreateObject("Microsoft.XMLHTTP")
   t = Timer
   objHTTP.Open "GET", "http://<%=Request.ServerVariables("SERVER_NAME")%>" & hiddenProfileURL.Value & "?" & txtDataURL.Value, False
   objHTTP.Send
   t = (Timer - t)
   size = Len(objHTTP.ResponseText)
   Set objHTTP = Nothing
   divDownloadTime.innerText = "Actual Time: " & t*256 & " cycles (" & CLng(t*1000) & " ms), " & size & " Bytes"
end function
-->
</SCRIPT>
</head>
<body onload="initSrc()">
<% WriteHeader
   If sCmd = "create" Then %>
  <font size=+1>
  Created profile file: <b><%=sProfileURL & "</b> (" & Round(Len(sNewData)/1024) & "K)" %><br>
  </font>
<% Else %>
  <font size=+1>Please copy the following source and save/upload it as:<br>
  &nbsp;&nbsp;http://<%=Request.ServerVariables("SERVER_NAME")%><b><%=sProfileURL%></b>
  <u>before</u> you click "Run" below.</font><br>
  <script language="VBSCRIPT">
  <!--
  Sub CopyToClipboard()
    Set f = document.frmOutput.taOutput
    f.focus
    f.select
    f.createTextRange.execCommand "Copy"
  End Sub
  -->
  </script>
  <form name="frmOutput" id="frmOutput">
  <textarea name="taOutput" id="taOutput" cols="60" rows="10" wrap="off">
  <%=Server.HTMLEncode(sNewData)%></textarea><br>
  <input type="button" id="btnCopy" name="btnCopy" value="Copy to Clipboard" onclick="CopyToClipboard()">
  </form>
<% End If %>
<p>

Profile URL: <%=sProfileURL%>?<input type="text" ID="txtDataURL" size=50 value=""><br>
<input type="hidden" ID="hiddenProfileURL" value="<%=sProfileURL%>">
<input type="button" ID="cmdRun" value="Run">
<input type="button" ID="cmdDownloadTime" value="Get Actual Time">
<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>">Profile Another File</a> <a href="#" onclick="window.open(hiddenProfileURL.Value + '?' + txtDataURL.Value)"><font size=1 color=#CCCCCC>View debug output</font></a><br>
<p>

<div id="divDownloadTime"></div>
<div id="divSummary"></div>
<div id="divResult"></div>

<p>
<% WriteFooter %>
</body>
</html>
