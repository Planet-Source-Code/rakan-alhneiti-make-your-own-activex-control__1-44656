<div align="center">

## make your own activeX control


</div>

### Description

my tutorial series 1 will help you to build an understanding for how activeX controls work and then you can make your own by me providing step by step.

its not mine but i got it from but i will provide all the series that u can learn from how to make activeX controls
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rakan Alhneiti](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rakan-alhneiti.md)
**Level**          |Intermediate
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rakan-alhneiti-make-your-own-activex-control__1-44656/archive/master.zip)





### Source Code

```
Good Morning and welcome to the first instalment in this ActiveX control tutorial.
I'm your amazingly geeky host Karl — and it's my job to ensure your ride on the Visual Basic train to ActiveX control land is an exciting one. Well, maybe just a slight-amusing one. Hmm, perhaps just a ride.
But don't let my pair of glass-bottle-bottom spectacles fool you - this isn't just a journey for mega-geeks.
Whatever Visual Basic programming experience you may have, learning about the wonderful world of ActiveX could improve your career, your bank balance and your love life*.
* (Love life claims based solely on life-long research by author Karl Moore and his numerous worldwide cyber-girlfriends)
Today, we'll be:
Getting the low-down on ActiveX components
Finding out the difference between ActiveX components and ActiveX controls
Discussing a few nerdy things (too boring to mention)
And... <drum roll> ... we'll be creating our own ActiveX control!
I know you're excited — but please, hold it in.
So without further ado, let's tootle off into the magical realms of ActiveX...
What's the difference between a horse trainer and a tailor? One tends a mare and one mends a tear!
And now for another belly-chuckler - what's the difference between an ActiveX component and an ActiveX control?
OK, not quite a dinner party puzzler but still an important point. Let's take a look at what exactly an ActiveX component is... and is not.
An ActiveX component is just a general term, encompassing:
An ActiveX EXE
An ActiveX DLL
An ActiveX Control
An ActiveX component is not:
Active, in any way, shape or form
A source of fibre that can help you lose weight as part of a calorie controlled diet
So what exactly are ActiveX EXEs and DLLs? Basically, they're chunks of code you use in your Visual Basic projects just by setting a reference to them — a little like how you set a reference to DAO or ADO when you need access to a database.
But that's another department completely... more to the point, just what are ActiveX controls?
Well, you might not know this, but you already have experience of ActiveX controls. You've used them, tweaked them and tossed them to one side — all in the course of a days work. Ohhhh yes.
Indeed, every time you set the Text property of a Text Box, you're utilising an ActiveX control. Every time you respond to the Click event of a Command Button, you're utilising an ActiveX control. Every time you run the MoveNext method of the Data control, you're utilising an ActiveX control.
I think you get the picture. In essence, an ActiveX control is anything you might see in the Toolbox.
Top Tip: Don't forget that you can also add more controls to the Toolbox by selecting Projects, Components
But how can all this background information help in everyday programming life?
Well, with the advent of Visual Basic 5 and, more recently Visual Basic 6, supercool geeky-types have been able to create their very own ActiveX controls.
So perhaps you could create your own groovy text box control that only allows the user to input numbers. Or perhaps just text. Or perhaps text and numbers, but no spaces.
Maybe you'd like to create a company-wide Exit button that flashes every time you hover your mouse over it. Sure, it might be about as useful as a pencil sharpener in the bullring, but it'd look good.
Other slightly more practical uses include creating a standardised Save dialog box. Or a lighter-weight version of the MSChart control. Or a plain but simple replacement for the InputBox() function. Or perhaps an intelligent scrollable window that displays a picture you pass it. Or a new and improved combo box. Or maybe just something else.
Then, when you need to use that groovy flashing Exit button, you simply draw it onto your form, just as you would any standard control. You could then set its MyControl.Forecolor property, and perhaps respond to its MyControl_Click event by adding a bit of code. You could even execute one of its' methods every now and then, such as MyControl.FlashAnimation.
The difference here is that you, as a productive, presentable, professional, pragmatic programmer, created the control. And as such, you dictate when the MyControl_Click event fires. Or how the MyControl.FlashAnimation method works. Or in which way the MyControl.ForeColor property is implemented — is the user presented with a text list of just four colours or the standard colour selection panel?
We'll be covering all this and more in this series. But for now, let's jump in at the deep end and knock out our very first ActiveX control!
Now, brace yourself as we prepare to create our own ActiveX control.
We're going to create a little option button that flashes a few times when you run a certain method. It's not overly useful, but could help grab a user's attention.
Start Visual Basic
Create a New 'ActiveX Control' project
A grey box should appear on your screen. This is your workspace — it's basically a form without a border, caption or minimize/maximise/close buttons.
And that makes sense, after all when did you last use a control that has its own close button?
First off, let's rename our ActiveX control:
Change the Name property of UserControl1 to 'Flasher'
Now change the Name property of Project1 to 'Animation'
Excellent! Now...
Double-click on the Option Button control in the toolbox
Remove the Caption property of the Option Button
Change its Name property to 'optFlasher'
We've just added an Option Button to the workspace. Now let's add the Timer control:
Double-click on the Timer control
Change its Name property to 'tmrAnimation'
That's great. Now I want you to resize a few of the things on your screen. We'll be doing all this resizing in code later on, but for now:
Move the Option Button to the very top left, so it just touches the corner edges of your workspace like this:
Now resize the workspace so it just touches the bottom edges of your Option Button like this:
Now the stuff you currently see in your workspace will become your 'control', the thing your user sees when adding it to their forms.
Hmm, it's about time we added some code. Not much, just a lil'.
Enter the code window by selecting View, Code
Type in the following code:
Public Sub Flash()
  tmrAnimation.Interval = 300
End Sub
This just sets the Interval property of tmrAnimation to around a third-of-a-second (300 milliseconds). When the Timer springs into action every 300-milliseconds, it fires its Timer event.
So let's add code to that...
In the Object drop-down list (which currently says General), select 'tmrAnimation'
The Procedure drop-down list next to it should say 'Timer' — if not, select the 'Timer' event from the list
Your screen should look a little like this at the moment:
Tap in the following code:
Static NoOfFlashes As Integer
  ' This is just a variable that holds
  ' a number - the 'Static' prefix just
  ' means it doesn't forget its value
  ' when this procedure is over...
  optFlasher.Value = Not (optFlasher.Value)
  ' Sets the value of our Option Button
  ' to the opposite of its current value...
  ' so if it's "on", it'll be turned off -
  ' and vice versa
  NoOfFlashes = NoOfFlashes + 1
  ' Increment the variable to show number
  ' of times we have "flashed"
  If NoOfFlashes = 8 Then
    ' If we've had eight separate flashes so far
    NoOfFlashes = 0
    ' Reset the NoOfFlashes...
    tmrAnimation.Interval = 0
    ' ... and turn off the timer
  End If
That's it! You've completed the creation of your first ActiveX control.
Now let's put it to the test...
Let's see what all that hard work has given us.
Click File, Add Project
Select 'Standard EXE' and click Open
Now we have two different projects open at the same time; our control and this new Standard EXE thing we've just created.
Let's add our new control to the Standard EXE now.
Drag out the Flasher control () on your toolbar onto Form1, like this:
Top Tip: If your Flasher control is greyed-out... it means your copy of Visual Basic has been attacked by huge killer bees from the terrifying jungles of Outer Mongolia or you haven't closed the workspace of your control. Hmm, probably the latter actually. Close the workspace and try again!
See how your Option Button appears?
Look in the Properties window. Can you see all the Properties your control already has? A Name property, TabIndex, ToolTipText... and more! These are all assigned by default.
Now...
Add a Command Button to the form
Place the following code behind it:
Flasher1.Flash
The method you've just tapped in is the one we coded!
When we added the 'Public Sub Flash' code, it's automatically turned into one heckuva groovy method!
Try hitting F5 and running your application. Now hit the Command Button! See what happens?
The Option Button should flash for a few seconds... great for highlighting a warning of some sort. But not so great at anything else.
This week, we've taken a brief tour of ActiveX controls. We found out exactly what they are and how they fit into the world of ActiveX components.
We even created our own basic, if not slightly useless ActiveX control!
Next week, we'll be getting even geekier; we'll be learning more about creating our own Methods... as well as covering Properties and Events.
So don't miss it... out next week at a newsagent near you.
But until then, this is your fantabulous host, Karl Moore, saying goodnight for tonight. Goodnight!
```

