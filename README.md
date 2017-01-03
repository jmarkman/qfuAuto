# Quote Follow Up Automation with Python

This is a GUI automation script for a mass emailing process at work; it removes the need for the user to completely supervise and go through the motions of running the query and shoving it through the Mail Merge process provided by Microsoft Word. We don't have access to our IT infrastructure at all and we're barred from using SMTP to send out emails programmatically. 

Find included an extra .py file that goes through the code step by step.

# Future
I'm trying to deprecate this script in favor of one that can make the database query all on its own and send out the email programmatically, but that requires a meeting with the IT guys across the hall, and that might happen either next week or within the next 6 months because it's "expensive" (it's actually high-ranking office politics between three people since there are fragile egos at stake when it comes to asking people for things up there).

*11/11/2016* - This meeting still hasn't happened :confused:

# Updates
*11/11/2016*

- I pried a chunk of it loose from hardcoding:

The worst part about this is it took me a few months without thinking about the script to come back to it and realize I could use os.path.expanduser to get the userpath and just append "\\Documents\\blah" to access the user's Documents folder. However, I created a new problem for myself.

*1/3/2017*
There's still no meeting. Instead I got a new query that's even bigger. I'm either going to have to learn about using COM in Python or switch this over to C# and use the Microsoft Office .dlls to access Mail Merge that way.

# Script issues
~~*9/26/2016*~~
*11/11/2016*

- I have to check that the folder "Quote Follow Ups Archive" exists for the user:

This shouldn't be too difficult. Getting the time to actually implement it, however, might be.
  
- I use time.sleep() a *lot*

This is another annoying issue that comes with the script; since this isn't sitting on a computer that just sits powered on in a corner 24/7, restarts make program startup times slow (no magic of SSDs here) and each program is pretty finnicky about its startup time as well as when it's "ready" to perform a function as simple as opening a file. Looking at it again from update 2, I guess it's acceptable but I don't have anyone in the office I can really turn to unless I go ask some random on the internet for review. In turn...
   
- It's slow

Python isn't a fast language, but I'm not really doing anything that requires extreme speed. Regardless, it takes longer than I feel it should to perform this task with all of the time.sleep() and general waiting for programs to complete their function. Still really hope I can deprecate this script. It was fun to work on initially but the more I use it, the more I wish I could have something better and easier.