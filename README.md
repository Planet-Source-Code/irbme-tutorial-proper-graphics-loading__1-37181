﻿<div align="center">

## Tutorial \- PROPER Graphics Loading


</div>

### Description

This is a tutorial which will show you how to load bitmaps, gifs, jpegs or whatever from a resource file into main memory and use them. Remember to download the ZIP it has some source you can use and an example. Also please vote and and give feedback.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-23 11:59:42
**By**             |[IRBMe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irbme.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Tutorial\_\-1095067232002\.zip](https://github.com/Planet-Source-Code/irbme-tutorial-proper-graphics-loading__1-37181/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./Tutorial_files/filelist.xml">
<title>Tutorial – Graphics Loading</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>Christopher Waddell</o:Author>
 <o:LastAuthor>Christopher Waddell</o:LastAuthor>
 <o:Revision>5</o:Revision>
 <o:TotalTime>76</o:TotalTime>
 <o:Created>2002-07-21T18:46:00Z</o:Created>
 <o:LastSaved>2002-07-21T20:40:00Z</o:LastSaved>
 <o:Pages>4</o:Pages>
 <o:Words>1369</o:Words>
 <o:Characters>7807</o:Characters>
 <o:Company>Developement</o:Company>
 <o:Lines>65</o:Lines>
 <o:Paragraphs>15</o:Paragraphs>
 <o:CharactersWithSpaces>9587</o:CharactersWithSpaces>
 <o:Version>9.4402</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:16.0pt;
	font-family:Arial;
	mso-font-kerning:16.0pt;}
h2
	{mso-style-next:Normal;
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:12.0pt;
	font-family:"Times New Roman";}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{margin:0cm;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:16.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	font-weight:bold;}
p.MsoSubtitle, li.MsoSubtitle, div.MsoSubtitle
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	font-weight:bold;}
@page Section1
	{size:595.3pt 841.9pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-GB style='tab-interval:36.0pt'>
<div class=Section1>
<p class=MsoTitle>Tutorial – Graphics Loading</p>
<p class=MsoNormal><b><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoSubtitle>PART 1 – The Resource File</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I have seen many ways of games storing and loading their
graphics. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>One of the most common ways is to store all the graphics in
separate files and load them all at runtime using the LoadPicture method. The
trouble with this is that in a finished game, anybody can just steal or change
the graphics with a simple bitmap editor like PAINT.<span style="mso-spacerun:
yes">  </span>This also applies to graphics which are all stored in the same
bitmap file.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Another common way, and probably the worst, is to store all
the graphics on a hidden form in picture boxes. This is probably the easiest
way but definitely the worst. It makes your compiled exe file huge, it is very
difficult to edit the graphics without a whole recompilation of the whole
project, and it takes up a large amount of memory (lots of picture boxes),
unless all the graphics are stored in one picturebox.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The best way, but unfortunately, not the most common, is to
store them all in a resource file. This means that it isn’t so easy for other
people to steal or change your graphics, you can easily update your graphics
without having to recompile your project, you’re compiled exe isn’t huge, and
you don’t have about 50 files, one for each graphic.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So the conclusion is: resource files are your friend! Use
them, and use them often. You don’t just have to store your graphics in them;
you can store text, graphics, sounds and anything else you want all in one
file.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So how do you use a resource file. Well, first get a
resource editor, VB has an inbuilt one and this is the one I use.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Click the add-ins menu, click add-in manager and load “VB 6
Resource Editor”. Press OK and now you should see a green icon on your toolbar.
If not, click the Tools Menu and it should be at the bottom. Click this once
and the VB6 resource editor should pop-up. On the resource editor’s toolbar,
press the “Add Custom” button, it looks like 4 silver squares. Here you can add
any file you want, we will store our graphics in here. Note, there is a button
to add graphics, but we will need to load them into the “custom” section, you
will see why later.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You should now see an open file dialog box. Find your
graphic and load it, now you should be able to name it. If it is the first one
you are loading it will be by default called “101” and I will keep it at that
for ease but you should name it properly to something you will remember. You
should also give your file a prefix so you remember its type.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>E.g.</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>JPG_Plane</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>BMP_Bullet</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>WAV_GunShot</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>MP3_BackgroundMusic</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I now have one JPEG file loaded into the custom section and
called “101” in my file. Once you have added your resources, save the resource
file with the same name as your project (e.g. “Project1.RES”) and now close the
editor.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now, how do we load our graphics from a resource file into a
Picturebox? Well, if you just want to do this, you should load your resources
into the picture section of your resource file and use:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Picture1.Picture = LoadResPicture(ID,Type)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>But as you will now see, we don’t want to load pictures into
pictureboxes for games, and also why we put our graphic in the custom section.
Ok, we first have to look at a new function to load the custom graphic.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private</span> <span
style='color:navy'>Sub</span> ExtractResource(FileName <span style='color:navy'>As</span>
<span style='color:navy'>String</span>, ResourceName <span style='color:navy'>As
Variant</span>)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">  </span><span
style='color:navy'>Dim</span> Buffer() <span style='color:navy'>As Byte</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>Buffer =
LoadResData(ResourceName, &quot;CUSTOM&quot;)</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>Open</span> FileName <span style='color:navy'>For Binary As</span>
#1</p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span><span
style='color:navy'>Put</span> #1, , Buffer</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>Close</span> #1</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>Erase</span> Buffer</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal><span style='color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>I wrote this function just there but it probably has many
other versions on the Internet. What it does is loads the resource into a byte
buffer and then loads that byte buffer into a file.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You call it like:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Call</span> ExtractResource
(“C:\MyGraphic.Jpg”,”101”)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>It is useful to note, this function will also work for any
other custom resource.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now all games (I should hope) will use Bitmaps, but these
make your resource files very large, so I store them in the resource files as
JPEG’s. Using this function and the resource file I made earlier, I now have a
file in C:\ called MyGraphic.Jpg. JPG’s are no good though, so I want to
convert it to a bitmap. It took me less than 4 minutes to solve this problem,
and here’s the result:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Sub</span>
ConvertToBitmap(FileName <span style='color:navy'>As String</span>, Extension <span
style='color:navy'>As String</span>)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">  </span><span
style='color:navy'>Dim</span> Pic <span style='color:navy'>As</span>
IPictureDisp</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>Set</span> Pic = <span style='color:navy'>New</span>
StdPicture</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>Set</span> Pic = LoadPicture(FileName &amp; &quot;.&quot;
&amp; Extension)</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>Kill FileName
&amp; &quot;.&quot; &amp; Extension</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>SavePicture Pic,
FileName &amp; &quot;.bmp&quot;</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>This is how to use it:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Call</span>
ConvertToBitmap(“C:\MyGraphic”,”JPG”)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>This will turn MyGraphic.Jpg into MyGraphic.Bmp. Nice huh!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Ok, I have now extracted my JPEG graphic from a resource
file and turned into a BMP file on the harddrive. Now we have something that’s
useable. We’re half way there now.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>PART 2 – LOADING THE GRAPHIC</h2>
<p class=MsoNormal><b><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></b></p>
<p class=MsoNormal>I’m going to paste a whole lot of commented code and
hopefully that should do it.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
GetDesktopWindow <span style='color:navy'>Lib</span> &quot;user32&quot; () <span
style='color:navy'>As Long</span></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
CreateCompatibleDC <span style='color:navy'>Lib</span> &quot;gdi32&quot; (<span
style='color:navy'>ByVal</span> hdc <span style='color:navy'>As Long</span>) As
<span style='color:navy'>Long</span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
LoadImage <span style='color:navy'>Lib</span> &quot;user32&quot; <span
style='color:navy'>Alias</span> &quot;LoadImageA&quot; (<span style='color:
navy'>ByVal</span> hInst <span style='color:navy'>As Long</span>, <span
style='color:navy'>ByVal</span> lpsz <span style='color:navy'>As String</span>,
<span style='color:navy'>ByVal</span> un1 <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> n1 <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> n2 <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> un2 <span style='color:navy'>As Long</span>)
<span style='color:navy'>As Long</span></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
SelectObject <span style='color:navy'>Lib</span> &quot;gdi32&quot; (<span
style='color:navy'>ByVal</span> hdc <span style='color:navy'>As Long</span>, <span
style='color:navy'>ByVal</span> hObject <span style='color:navy'>As Long</span>)
<span style='color:navy'>As Long</span></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
GetObject <span style='color:navy'>Lib</span> &quot;gdi32&quot; <span
style='color:navy'>Alias</span> &quot;GetObjectA&quot; (<span style='color:
navy'>ByVal</span> hObject <span style='color:navy'>As Long</span>, <span
style='color:navy'>ByVal</span> nCount <span style='color:navy'>As Long</span>,
lpObject <span style='color:navy'>As Any</span>) <span style='color:navy'>As
Long</span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
GetDC <span style='color:navy'>Lib</span> &quot;user32&quot; (<span
style='color:navy'>ByVal</span> hwnd <span style='color:navy'>As Long</span>) <span
style='color:navy'>As Long</span></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
DeleteObject <span style='color:navy'>Lib</span> &quot;gdi32&quot; (<span
style='color:navy'>ByVal</span> hObject <span style='color:navy'>As Long</span>)
<span style='color:navy'>As Long<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'>Private Declare Function</span>
DeleteDC <span style='color:navy'>Lib</span> &quot;gdi32&quot; (<span
style='color:navy'>ByVal</span> hdc <span style='color:navy'>As Long</span>) <span
style='color:navy'>As Long</span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Type</span> BITMAP</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmType<span
style="mso-spacerun: yes">              </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">      </span><span style='color:green'>'Bitmap type</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmWidth<span
style="mso-spacerun: yes">            </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">    </span><span style="mso-spacerun: yes">  </span><span
style='color:green'>'Width/Pixels<o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmHeight<span
style="mso-spacerun: yes">           </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">      </span><span style='color:green'>'Height/Pixels</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmWidthBytes<span
style="mso-spacerun: yes">   </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">      </span><span style='color:green'>'Width/Bytes</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmPlanes<span
style="mso-spacerun: yes">           </span><span style='color:navy'>As Integer</span><span
style="mso-spacerun: yes">    </span><span style='color:green'>'Planes</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmBitsPixel<span
style="mso-spacerun: yes">       </span><span style='color:navy'>As Integer</span><span
style="mso-spacerun: yes">    </span><span style='color:green'>'Bits per Pixel</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>bmBits<span
style="mso-spacerun: yes">               </span><span style='color:navy'>As
Long</span><span style="mso-spacerun: yes">       </span><span
style='color:green'>'Bits<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'>End Type<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Public Type</span> GRAPHIC</p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>hGraphic<span
style="mso-spacerun: yes">       </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">     </span><span style='color:green'>'Handle</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>hdc<span
style="mso-spacerun: yes">                </span><span style='color:navy'>As
Long</span><span style="mso-spacerun: yes">     </span><span style='color:green'>'Device
context</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>hWidth<span
style="mso-spacerun: yes">          </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">     </span><span style='color:green'>'Width/Pixels</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>hHeight<span
style="mso-spacerun: yes">         </span><span style='color:navy'>As Long</span><span
style="mso-spacerun: yes">     </span><span style='color:green'>'Height/Pixels</span></p>
<p class=MsoNormal><span style='color:navy'>End Type<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Const</span>
LR_LOADFROMFILE<span style="mso-spacerun: yes">  </span><span style='color:
navy'>As Integer</span> = &amp;H10</p>
<p class=MsoNormal><span style='color:navy'>Private Const</span>
IMAGE_BITMAP<span style="mso-spacerun: yes">     </span><span style='color:
navy'>As Integer</span> = 0</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Public Sub</span> DeleteGraphic(<span
style='color:navy'>ByRef</span> udtGraphic <span style='color:navy'>As</span>
GRAPHIC)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>DeleteDC
udtGraphic.hdc<span style="mso-spacerun: yes">                 </span><span
style='color:green'>'Delete the Device Context</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span>DeleteObject
udtGraphic.hGraphic<span style="mso-spacerun: yes">        </span><span
style='color:green'>'Delete the object<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Public Function</span> LoadBitmap(<span
style='color:navy'>ByVal</span> Path <span style='color:navy'>As String</span>)
<span style='color:navy'>As</span> GRAPHIC</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">  </span><span
style='color:navy'>Dim</span> hDesktop <span style='color:navy'>As Long</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">  </span><span
style='color:navy'>Dim</span> bmBitmap <span style='color:navy'>As</span>
BITMAP</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:green'><span style="mso-spacerun:
yes">    </span>'Load the bitmap, this will return a handle if successful<o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">   
</span>LoadBitmap.hGraphic = LoadImage(App.hInstance, Path, IMAGE_BITMAP, 0, 0,
LR_LOADFROMFILE)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>If</span> LoadBitmap.hGraphic = 0 <span style='color:navy'>Then</span><span
style="mso-spacerun: yes">     </span><span style='color:green'>'If no handle,
then the loadimage function must have failed _<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:green'><span style="mso-spacerun:
yes">       </span>Most probable cause is because it dousn't exist or the given
path is wrong</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span>Err.Raise _</p>
<p class=MsoNormal><span style="mso-spacerun: yes">                 
</span>Number:=vbObjectError + 513, _</p>
<p class=MsoNormal><span style="mso-spacerun: yes">                 
</span>Source:=&quot;LoadBitmap(''&quot; &amp; Path &amp; &quot;'')&quot;, _</p>
<p class=MsoNormal><span style="mso-spacerun: yes">                 
</span>Description:=(&quot;The graphic cannot be loaded. Please ensure that
''&quot; &amp; Path &amp; &quot;'' exists, as this is the most probable
cause.&quot;)</p>
<p class=MsoNormal><span style="mso-spacerun: yes">      </span><span
style='color:navy'>Else</span><span style="mso-spacerun:
yes">                       </span><span style="mso-spacerun:
yes">                                    </span><span style='color:green'>'We
got a handle, the bitmap's loaded. So...<o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span>hDesktop =
GetDesktopWindow<span style="mso-spacerun:
yes">                                 </span><span style='color:green'>'Get
desktop handle</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">       
</span>LoadBitmap.hdc = CreateCompatibleDC(GetDC(hDesktop))<span
style="mso-spacerun: yes">        </span><span style='color:green'>'Create a
new DC compatible with the desktop<o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span>SelectObject
LoadBitmap.hdc, LoadBitmap.hGraphic<span style="mso-spacerun: yes">           
</span><span style='color:green'>'Select the graphic, and the new DC into an
object<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span>GetObject
LoadBitmap.hGraphic, Len(bmBitmap), bmBitmap<span style="mso-spacerun:
yes">      </span><span style='color:green'>'Get the width and height of the
graphic</span></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span><span
style="mso-spacerun: yes">       </span>LoadBitmap.hWidth = bmBitmap.bmWidth</p>
<p class=MsoNormal><span style="mso-spacerun: yes">       
</span>LoadBitmap.hHeight = bmBitmap.bmHeight</p>
<p class=MsoNormal><span style="mso-spacerun: yes">        </span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes">    </span><span
style='color:navy'>End If<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>End Function<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>Ok, just in case you’re wondering, I again wrote that.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Just use like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Dim</span> bmBitmap<span
style="mso-spacerun: yes">        </span><span style='color:navy'>As</span>
GRAPHIC</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>BmBitmap = LoadBitmap(“C:\MyGraphic.Bmp”)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:green'>‘Do stuff with it<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>DeleteGraphic bmBitmap</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You now have a UDT called bmBitmap (or whatever) which
contains a handle and a Device Context to your graphic, and also its width and
height in pixels. This should be all the information you need to draw your
graphic using BitBlt, StretchBlt, TransparentBlt, PatBlt and any other windows
GDI functions that require DC’s, Width’s or Height’s.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I must emphasise to you to remember to use DeleteGraphic
once you are finished, otherwise you will have a memory leak. And only one
thing left: delete the bitmap that was created on the hard-drive.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Kill “C:\MyGraphic.Bmp”</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>OK, I think that’s it.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Hopefully you should have learned how to:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Put graphics into a RES file</p>
<p class=MsoNormal>Load the graphics from the RES file into a file on the
hard-drive</p>
<p class=MsoNormal>Convert the file to a bitmap (if needed)</p>
<p class=MsoNormal>Load the file into memory with useful information like DC</p>
<p class=MsoNormal>Delete the Graphic</p>
<p class=MsoNormal>Delete the File</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>If you download the ZIP you will find the offline html version
of this tutorial, and a useful stand-alone module containing all the useful
functions mentioned above + a project showing how everything in this project
comes together by using the module.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Please remember to vote and leave feedback. And I expect to
see this in your future games (Just mention my name in the credits if you use
my module by the way).</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

