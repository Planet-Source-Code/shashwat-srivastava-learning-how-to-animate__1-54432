<div align="center">

## Learning How to Animate


</div>

### Description

Learn easily how to animate in VB with simple controls. No API nothing!!! And please do vote so that I can write more projects and articles :-)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-06-18 08:23:34
**By**             |[Shashwat Srivastava](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shashwat-srivastava.md)
**Level**          |Beginner
**User Rating**    |3.3 (20 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Learning\_H1758626172004\.zip](https://github.com/Planet-Source-Code/shashwat-srivastava-learning-how-to-animate__1-54432/archive/master.zip)





### Source Code

<html><div style='background-color:'><DIV class=RTE></DIV>
<H1 class=RTE align=center>Learning How To Animate</H1>
<P class=RTE align=left>If you know nothing about visual basic and want to create animation in visual basic this is the right place where you have landed. So move ahead.</P>
<P class=RTE align=center><STRONG>Contents</STRONG></P>
<OL>
<LI>
<DIV class=RTE align=left>Chapter 1 - Introduction to Visual Basic</DIV></LI>
<LI>
<DIV class=RTE align=left>Chapter 2 - Introduction to Picture Box, Timer and Variable</DIV></LI>
<LI class=RTE>Chapter 3 - Animating</LI></OL>
<H1 align=center>Chapter-1</H1>
<H1 align=center>Introduction to Visual Basic</H1>
<P>Visual Basic is an environment in which you can create applications. Infact its an application producing machine. Its not merely a language. And with the coming up of Visual Basic 6.0 it has become all the more interactive. Internet features have been added.</P>
<P>Friend I am not going to introduce Visual Basic in more detail. When you will use it then you will come to know about it in a better way. So lets move ahead to next chapter.</P>
<H1 align=center>Chapter-2</H1>
<H1 align=center>Introduction to Picture Box, Timer and Variable</H1>
<P align=left>So lets start creating our project. Follow the following steps - </P>
<OL>
<LI>
<DIV align=left> Start Visual Basic 6.0 by selecting it from Start menu.</DIV></LI>
<LI>
<DIV align=left> Visual Basic will prompt you to create a new project. If it doesn't then click on File menu and select New Project.</DIV></LI>
<LI>
<DIV align=left> New Project dialog box will be displayed. Select "Standard EXE" and click on OK.</DIV></LI></OL>
<P align=left>Change the properties of the form to the following - </P>
<P>
<TABLE width="75%" border=1>
<TBODY>
<TR>
<TD>Name</TD>
<TD>
<P>frmanimation</P></TD></TR>
<TR>
<TD>BackColor</TD>
<TD>
<P>From the popup menu choose the black color</P></TD></TR>
<TR>
<TD>Caption</TD>
<TD>Animation</TD></TR>
<TR>
<TD>ScaleMode</TD>
<TD>3 - Pixel</TD></TR></TBODY></TABLE></P>
<P align=left>What is picture box?</P>
<P align=left>In our project we want to animate a picture, there are many ways to do it. But this one is the simplest. Just you think of a picture box, a box or control which can hold picture with certain properties.</P>
<P align=left>Double click on the picture box icon. After doing this change the properties of the Picture Box to the following - </P>
<P>
<TABLE width="75%" border=1>
<TBODY>
<TR>
<TD>Name</TD>
<TD>
<P>picanimation</P></TD></TR>
<TR>
<TD>AutoRedraw</TD>
<TD>
<P>True</P></TD></TR>
<TR>
<TD>AutoSize</TD>
<TD>True</TD></TR>
<TR>
<TD>BackColor</TD>
<TD>From the popup menu choose the black color</TD></TR></TBODY></TABLE></P>
<P align=left>Now what we wan to do is to load a picture (the picture should not be too big). There are two ways to do this job. I will discuss both. First one which I generally do not use is that one can load picture by selecting Picture property of your Picture Box and browsing and then loading the picture.</P>
<P align=left>In this process there is very serious problem. For example if the your picture is in My Documents folder, then the path of the picture will be "C:\My Documents\xyz.bmp". If you decide to email your program to one of your friend and lets suppose that he stores this program in the some other folder, then the program won't be able to load the picture.</P>
<P align=left>To solve this problem there is a solution as you know "To every problem there exists a solution". In the Form_Load event type the following - </P>
<P align=left><CODE><I><B>picanimation.Picture = LoadPicture(App.Path & "\xyz.bmp")</B></I></CODE></P>
<P align=left>Now let me explain every thing clearly. First of all you might be thinking what is this Form_Load event but it is clear form its name it event that is executed as soon the form loads.</P>
<P align=left>Now what about the code? Its very simple.</P>
<P align=left>Here picanimation is the name of the picture. Picture is its one of the property. What we simply do is that we load picture by giving LoadPicture command. I have then typed App.path, this means application's path and then we add the name of the picture with a slash.</P>
<P align=left>Now its time to animate!</P>
<P align=left>One thing we need is Timer. This is another control which you can easily locate an Toolbar.</P>
<P align=left>Change the properties of the timer to following - </P>
<P>
<TABLE width="75%" border=1>
<TBODY>
<TR>
<TD>Name</TD>
<TD>
<P>tmranimation</P></TD></TR>
<TR>
<TD>Enabled</TD>
<TD>
<P>True</P></TD></TR>
<TR>
<TD>Interval</TD>
<TD>1</TD></TR></TBODY></TABLE></P>
<P align=left><BR>Now its time to do the real job so move on to the next chapter where I will explain you about the timer in much more detail.</P>
<H1 align=center>Chapter-3</H1>
<H1 align=center>Animating</H1>
<P>For understanding timer you think of timer to be like a clock. It will keep on doing the work specified in the _timer( ) event after every definite interval specified. For example in this project we have set the interval 1, so after every 1 millisecond the command will be executed which is entered in _timer( ) event.</P>
<P>Now lets do the real job which we are waiting for. First of enter the following code in the Form_Load event</P>
<P><B><I><CODE>picanimation.Picture = LoadPicture(App.Path & "\xyz.bmp")</CODE></I></B></P>
<P>where 'xyz.bmp' is the name of the picture. Remember before entering this code your project must be saved and picture should be in the same folder in which the project is saved.</P>
<P>Now lets come on to the code of timer. Enter the following code in the tmranimation_timer( ) event.</P>
<P><B><I><CODE>If frmanimation.ScaleHeight > picanimation.Top Then</CODE></I></B></P>
<P><CODE><I><B>picanimation.Top = picanimation.Top + 10 </B></I></CODE></P>
<P><CODE><I><B>Else picanimation.Top = 1 </B></I></CODE></P>
<P><CODE><I><B>End If</B></I></CODE></P>
<P>Now you may be wondering what this a tiny bit of code does? But infact it does the real job. Let me translate the code in simple english. There is condition that if the scaleheight of the form is greater than the top of the picture then the picture's top is increased means that the picture moves down and if disappears from the form then it is again restored to the beginning. That's all. </P>
<P>Happy Programming</P>
<P>By Shashwat Srivastava.</P></div></html>

