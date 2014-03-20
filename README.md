myToolBox
=========

This is an application I wrote while working as a desktop technician and is filled with tools that made my job much easier. Since a lot of the tasks were client specific and specific to their network, I am currently rewriting the applicaiton to work anywhere. Also, the entire application has a whole list of files I've yet to commit. I'm currently working on stripping any of my job specific stuff out of them and I will upload them as I have time.

If you want to use this application for yourself just know that I am still working on making the changes to the code necessary to work outside of my current job, but it should have most of the features working at this point. Also, copy the "Phrasekeys" folder and its contents to your C:\ (root). This is obviously only designed to run on Windows XP or higher. I have not tested it on Windows 8 yet.

Toolbox.hta is the main application, Phrasekeys.hta is just the Phrasekeys section of the main app broken off into a standalone because some of my peers requested it, and AdminRightsTool.hta is a little tool I broke off of the main Toolbox app for manipulating admin rights on remote machines.

I've added importUsageReport.vbs. I use it daily to import the usage statistics to an Excel spreadsheet to track how useful this app has been for my peers and I. Every Monday it also fills some data in on a YTD soreadsheet tracking my ticket count. No one is going to find this file useful as it is, but this can be used as an outline to create your own script. Plus, having it here means I have a backup stored somewhere.

I've also added ListConverter.hta and its functionality now exists in the main Toolbox.hta as well. This will take a vertical list and change it to a comma separated list with the option to add the word 'and' after the last comma; or it will do the reverse action and give you a vertical list. I saw a need for this when trying to import large lists of computer names from a text file I used to run some batch commands in toolbox.hta to a format appropriate for a ticket note. I found myself hitting comma, delete, space after each computer name then added 'and' after the last comma and I figured that I do this often enough to warrant writing something to automate this as well.
