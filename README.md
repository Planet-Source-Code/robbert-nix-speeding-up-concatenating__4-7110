<div align="center">

## speeding up concatenating


</div>

### Description

Concatenating is not very efficient in VBS, especially when a large number of small string are concatenated. Here is how to speed it up
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robbert Nix](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robbert-nix.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robbert-nix-speeding-up-concatenating__4-7110/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./Concatening_files/filelist.xml">
<title>Speeding up String Concatenation</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>R. Nix</o:Author>
 <o:LastAuthor>R. Nix</o:LastAuthor>
 <o:Revision>2</o:Revision>
 <o:TotalTime>4</o:TotalTime>
 <o:Created>2002-01-08T13:10:00Z</o:Created>
 <o:LastSaved>2002-01-08T13:10:00Z</o:LastSaved>
 <o:Pages>3</o:Pages>
 <o:Words>504</o:Words>
 <o:Characters>2875</o:Characters>
 <o:Lines>23</o:Lines>
 <o:Paragraphs>5</o:Paragraphs>
 <o:CharactersWithSpaces>3530</o:CharactersWithSpaces>
 <o:Version>9.2720</o:Version>
 </o:DocumentProperties>
 <o:OfficeDocumentSettings>
 <o:DoNotRelyOnCSS/>
 </o:OfficeDocumentSettings>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
 <w:DocumentKind>DocumentEmail</w:DocumentKind>
 <w:DrawingGridHorizontalSpacing>0 pt</w:DrawingGridHorizontalSpacing>
 <w:DisplayHorizontalDrawingGridEvery>2</w:DisplayHorizontalDrawingGridEvery>
 <w:DisplayVerticalDrawingGridEvery>2</w:DisplayVerticalDrawingGridEvery>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:536871559 0 0 0 415 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Verdana;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:11.0pt;
	mso-bidi-font-size:16.0pt;
	font-family:Verdana;
	mso-bidi-font-family:Arial;
	mso-font-kerning:16.0pt;}
h2
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:10.0pt;
	mso-bidi-font-size:14.0pt;
	font-family:Verdana;
	mso-bidi-font-family:Arial;
	font-style:italic;}
h3
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:3;
	font-size:13.0pt;
	font-family:Arial;}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:144.0pt center 216.0pt right 432.0pt;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Verdana;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 216.0pt right 432.0pt;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Verdana;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
p.MsoAutoSig, li.MsoAutoSig, div.MsoAutoSig
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Verdana;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
span.EmailStyle17
	{mso-style-type:personal;
	mso-ansi-font-size:10.0pt;
	mso-ascii-font-family:Arial;
	mso-hansi-font-family:Arial;
	mso-bidi-font-family:Arial;
	color:black;}
@page Section1
	{size:595.45pt 841.7pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;
	mso-header-margin:36.0pt;
	mso-footer-margin:36.0pt;
	mso-header:url("./Concatening_files/header.htm") h1;
	mso-footer:url("./Concatening_files/header.htm") f1;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-GB link=blue vlink=purple style='tab-interval:36.0pt'>
<div class=Section1>
<h1><b style='mso-bidi-font-weight:normal'><font size=2 face=Verdana><span
style='font-size:11.0pt'><span style='mso-bidi-font-size:16.0pt'>Speeding up
String Concatenation</span></span></font></b></h1>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>When you dynamically want to generate an
output page using information from a table, you could use the following
function to perform this operation.</span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<h3><b style='mso-bidi-font-weight:normal'><font size=4 face=Arial><span
style='font-size:13.0pt'>Method 1</span></font></b></h3>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>Function TableContent(&#8230;)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim strContent<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>&#8230;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Do While Not recordset.EOF<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>StrContent = strContent &amp;
recordset.Fields([cell])<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>&#8230;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Loop<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>TableContent = strContent<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>End Function<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>Method 1 is not an efficient if the number of
concatenations are large, since the concatenating operation &lt;strContent =
strContent &amp; recordset.Fields([cell])&gt; copies strContent and
recordset.Fields([cell]) to a temporary space in addition to a copy of the
result. The total number of copied characters is in the order of the
result-string length times the number of concatenations.</span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>In VB this problem can be avoided by using
the mid$ function to replace a part of a long empty string
&lt;mid$(strContent,&#8230;)=[newstr]&gt;, but in VBS this mid function cannot be
used this way. </span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>In VBS the Method 2 can be used with the same
end result.<span style="mso-spacerun: yes">&nbsp; </span>In this code part(index)
will be at least 30*2^i bytes long. (Since strContent has 30 elements, the
function will work until the maximum length of a string is reached. Currently
this is over 2GB) </span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>It can be easily verified that every
character is part of a concatenation operation only max(index) times. </span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>Therefore, the total number of copied
characters due to concatenations will be in the order of the length of the
result string times the logarithm of the result sting length (strict
mathematical proof is difficult, since the length of the new-added strings may
vary). The total number of concatenations will be the same this way, but since
short string concatenations occur frequently, while long string concatenations
occur less often, the total length of all concatenated strings decreases.</span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<b><font size=4 face=Arial><span style='font-size:13.0pt;font-family:Arial;
mso-fareast-font-family:"Times New Roman";mso-ansi-language:EN-GB;mso-fareast-language:
EN-US;mso-bidi-language:AR-SA;font-weight:bold'><br clear=all style='page-break-before:
always'>
</span></font></b>
<h3><b style='mso-bidi-font-weight:normal'><font size=4 face=Arial><span
style='font-size:13.0pt'>Method 2</span></font></b></h3>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>Function TableContent(&#8230;)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim strContent(30)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>&#8230;<o:p></o:p></span></span></font></p>
<p class=MsoNormal style='text-indent:36.0pt'><font size=2 color=blue
face="Courier New"><span style='font-size:10.0pt;font-family:"Courier New";
color:blue'><span style='mso-bidi-font-size:12.0pt'>&#8216;</span></span></font><font
color=lime face="Courier New"><span style='font-family:"Courier New";
color:lime'>sub to clear the content string</span></font><font color=blue
face="Courier New"><span style='font-family:"Courier New";color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>ClearString strContent<span style='mso-tab-count:
1'>&nbsp; </span><o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Do While Not recordset.EOF<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></font><font
color=lime face="Courier New"><span style='font-family:"Courier New";
color:lime'>&#8216;add to the result string</span></font><font color=blue
face="Courier New"><span style='font-family:"Courier New";color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>AddString strContent,
recordset.Fields([cell])<span style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>&#8230;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Loop<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>TableContent = fnReadString(strContent)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>End Function<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>&#8216;</span></span></font><font color=lime
face="Courier New"><span style='font-family:"Courier New";color:lime'>The
following subs and functions need to be included</span></font><font color=blue
face="Courier New"><span style='font-family:"Courier New";color:blue'> <o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>Sub ClearString(part)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim index<span style='mso-tab-count:1'> </span>&#8216;as
integer<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>For index = 0 to 30<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>part(index) = &#8220;&#8221;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Next<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>End Sub<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>Sub AddString(part, newString)<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim tmp<span style='mso-tab-count:1'>&nbsp;&nbsp; </span>&#8216;as
string<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim index<span style='mso-tab-count:1'> </span>&#8216;as
integer<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>part(0) = part(0) &amp; newString<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>If Len(part(0)) &gt; 30 Then<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>index = 0 <o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>tmp = &#8220;&#8221;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Do <o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:2'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>tmp = part(index) &amp; tmp<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:2'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>part(index) = &#8220;&#8221;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:2'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>index = index + 1<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Loop until part(index) = &#8220;&#8221;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>part(index) = tmp<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>End If<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>End Sub<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<font size=2 color=blue face="Courier New"><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;font-family:"Courier New";mso-fareast-font-family:
"Times New Roman";color:blue;mso-ansi-language:EN-GB;mso-fareast-language:EN-US;
mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'>
</span></font>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'>Fu</span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'>nction fnReadString(part)<o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim tmp<span style='mso-tab-count:1'>&nbsp;&nbsp; </span>&#8216;as
string<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Dim index<span style='mso-tab-count:1'> </span>&#8216;as
integer<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>tmp = &#8220;&#8221;<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>For index = 0 to 30<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>If part(index) &lt;&gt; &#8220;&#8221; Then<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:2'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>tmp = part(index) &amp; tmp<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>End If<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>Next<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
style='mso-bidi-font-size:12.0pt'>FnReadString = tmp<o:p></o:p></span></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><![if !supportEmptyParas]>&nbsp;<![endif]></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 color=blue face="Courier New"><span
style='font-size:10.0pt;font-family:"Courier New";color:blue'><span
style='mso-bidi-font-size:12.0pt'>End Function</span></span></font><font
face="Courier New"><span style='font-family:"Courier New"'><o:p></o:p></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><b><font size=4 face=Arial><span style='font-size:13.0pt;
font-family:Arial;font-weight:bold'>Comparing both methods</span></font></b></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>Consider concatenating the alphabet one
character at the time. Then the following concatenations occur.</span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>For simplicity sake the maximum length of
part(0) in the AddString subroutine is set to 1 (</span></span></font><font
color=blue face="Courier New"><span style='font-family:"Courier New";
color:blue'>If Len(part(0) &gt; 1 Then)</span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><b><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt;font-weight:bold'>Step<o:p></o:p></span></font></b></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border:solid windowtext .5pt;
 border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><b><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt;font-weight:bold'>Concatenations method 1<o:p></o:p></span></font></b></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border:solid windowtext .5pt;
 border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><b><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt;font-weight:bold'>Concatenations method 2<o:p></o:p></span></font></b></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>1<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>2<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>a&amp;b<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>a&amp;b<o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>3<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>ab&amp;c<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>4<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abc&amp;d<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>c&amp;d, ab&amp;cd<o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>5<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcd&amp;e<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>6<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcde&amp;f<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>e&amp;f<o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>7<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcdef&amp;g<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>8<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcdefg&amp;h<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>g&amp;h, ef&amp;gh, abcd&amp;efgh<o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>9<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcdefgh&amp;I<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>10<o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>abcdefghi&amp;j<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>I&amp;j<o:p></o:p></span></font></p>
 </td>
 </tr>
 <tr>
 <td width=49 valign=top style='width:36.9pt;border:solid windowtext .5pt;
 border-top:none;mso-border-top-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 <td width=180 valign=top style='width:135.0pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><font size=1 face=Verdana><span style='font-size:8.0pt;
 mso-bidi-font-size:12.0pt'>etc.<o:p></o:p></span></font></p>
 </td>
 <td width=174 valign=top style='width:130.5pt;border-top:none;border-left:
 none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
 mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><font size=1
 face=Verdana><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></font></p>
 </td>
 </tr>
</table>
<p class=MsoHeader style='tab-stops:36.0pt'><font size=2 face=Verdana><span
style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><span
style='mso-bidi-font-size:12.0pt'>I have run tests, series of 10,000 string
concentrations with a length of 100 * (rnd() ^ 3). Using method 1, it took 30
seconds to process the script, while method 2 took less than a second. I used a
Pentium 3, 500 Mhz Processor with 256Mb RAM machine.</span></span></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
<p class=MsoNormal><font size=2 face=Verdana><span style='font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]></span><o:p></o:p></font></p>
</div>
</body>
</html>
```

