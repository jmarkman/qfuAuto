# Quote Follow Up Automation with Python

This is a GUI automation script for a mass emailing process at work; it removes the need for the user to completely supervise and go through the motions of running the query and shoving it through the Mail Merge process provided by Microsoft Word. We don't have access to our IT infrastructure at all and we're barred from using SMTP to send out emails programmatically. 

Find included an extra .py file that goes through the code step by step.

# Future
I'm trying to deprecate this script in favor of one that can make the database query all on its own and send out the email programmatically, but that requires a meeting with the IT guys across the hall, and that might happen either next week or within the next 6 months because it's "expensive" (it's actually high-ranking office politics between three people since there are fragile egos at stake when it comes to asking people for things up there).

# Script issues
*9/26/2016*

There's a lot that I don't like about this although it "works" and does the job its supposed to:
-It's hardcoded
   I did a lot of research in terms of breaking away from hardcoded filepaths, but it ended up being more troubles than it was worth to apply os.path commands to most of the filepaths for program files and Excel file storage. I'd also have to deal with the documents within, although I think I can overall improve this by shoving all of the necessary things into a folder like I did for all the elements.
   
   Not to mention that this is GUI-specific and works **only** on the primary monitor without issue. Each image in the elements folder is a screenshot of a button or UI element I want to focus on and send a click event to. This actually makes working with resized windows much easier but still isn't a great solution.
   
-I use time.sleep() a **lot**
   This is another annoying issue that comes with the script; since this isn't sitting on a computer that just sits powered on in a corner 24/7, restarts make program startup times slow (no magic of SSDs here) and each program is pretty finnicky about its startup time as well as when it's "ready" to perform a function as simple as opening a file. In turn...
   
-It's slow
   Python isn't a fast language, but I'm not really doing anything that requires extreme speed. Regardless, it takes longer than I feel it should to perform this task with all of the time.sleep() and general waiting for programs to complete their function. I really hope I can deprecate this script. It was fun to work on initially but the more I use it, the more I wish I could have something better and easier.