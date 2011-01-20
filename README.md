Modified Signature Survey
==========================

What use is it?
--------------------------

First I'll summarize a few of the neat points of the original [Signature Survey](http://c2.com/doc/SignatureSurvey/):

* Quick visual summary of code of a file by file basis.
* Easy to distinguish classes with low and high complexity 
* Helps represent a few defining characteristics of classes in an easy to remember form

And a few things that I think it could have done better that I ended up implementing in mine:

* Distinction between methods, if statements, and loops. 
* Would be really nice to be able to tell at a glance how big methods tend to be; whether there are lots of small ones or a few huge ones 
* Once there are distinctions between methods, if statements, and loops, then color should be used to 
    1. mitigate the amount of syntactic noise 
    2. make the quick visual summary even easier to grasp
    3. highlight possible trouble points
* Since we're parsing the whole file like this anyway, we may as well include a line count and method count in the 'signature' of a class.

And a few things that were adjusted because of language and personal preference:

* VB.NET doesn't use curly braces or semicolons in the same way as java/c#, alternatives will have to be used
* In the 'detailed view' of each file, I would prefer to see just a list of the methods, with parameter types and the return type, instead of the whole file. I can open the whole file in my main editor, but having a list of the methods in the order they are in is quite helpful. So this is what I have done in my implementation.


Examples
--------------------------

[(Click for full size)](http://i.imgur.com/KRRQl.png)

[![Picture of a VB.NET signature survey page](http://i.imgur.com/KRRQl.png)](http://i.imgur.com/KRRQl.png)
