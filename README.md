# excelswissknife
Free multi-tool Microsoft Excel Add-in

Summary
----------------------------
1. What is Excel Swiss Knife?
2. Installation/Update
3. User instructions
----------------------------


1. What is Excel Swiss Knife?

Excel Swiss Knife (ESK) is an add-in for Microsoft Excel.
Once installed, add a new tab to the Excel menu, from which you can 
start all its functions. These are divided into more categories, and aim 
to speed up processing tasks and data transformation either manually 
or through native functions of Excel would require a lot of time and 
numerous steps, or in some cases they would not be possible.
 
The choice of the implemented functions comes from a long experience
in the elaboration of data bases in Excel, and from the continuous 
repetition of the same tedious operations for cleaning and rearranging 
data. I have then decided to take the situation in my hand and create 
a set of tools that could make my daily work easier. Now I'm releasing 
these tools to everyone, hoping they can be useful to you as much 
as they are to me!

I will keep on developing new features to add to the program, e
I would be very happy to receive your suggestions, which you can send
directly to excelswissknife@gmail.com or by connecting
to the www.forumexcel.it forum (italian language), 
where ESK has a dedicated topic

ESK is free for personal use, and for now available in Italian and
English (the latter with possible inaccuracies).


2. Installation and update

Since this is basically a collection of macros in Visual Basic
for Applications (VBA), you must enable it before installation
execution of the macros in Excel, and accept any request for
activate them during installation. I obviously guarantee the absence of
any malicious code within the program. It is also necessary to
close all open Excel instances before proceeding.


3. User instructions

Every function of the program has its particularities, and goes beyond the
The purposes of this brief presentation to go down into the details of each one
of them. Guides with specifications for each function are published on
the extension's website http://www.excelswissknife.com.

Generally speaking, before using the functions of the program, I
stress that it is ALWAYS necessary to back up your data, as it is not 
technically possible to reverse the effect of a macro. 
Furthermore, in Excel the execution of a macro "destroys" the UNDO's stack, 
so in most of cases IT WILL NOT BE POSSIBLE TO RETURN TO THE STATE PREVIOUS 
TO THE EXECUTION OF THE FUNCTION.

To overcome, at least partially, this risk, all "destructive" functions
of the program provide for the creation of a backup sheet (named 
"ESK_Backup" and with the tab highlighted in yellow).
So if something goes wrong, it will be at least possible to get back to 
recover data before the last function was performed, taking it from the
backup sheet. However, the backup sheet is overwritten each
time, so if you perform multiple functions consecutively, it will be
always possible to go back ONLY BEFORE THE VERY LAST PROCESSING,
while the previous ones will be now irreversible, if not by closing the 
file without saving it (and thus losing any further modification
carried out).
