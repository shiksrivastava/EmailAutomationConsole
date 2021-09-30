# EmailAutomationConsole
Google Scripts Code to automate client acquisition sourcing emails, follow-ups, and check status

## Features
- send mass emails with variables in body per specified draft
- can add variables in subject as well
- send followups in thread per another specified draft
- reports whether email was invalid (bounced) or sent
- automatically checks and updates the status if the email has been replied to

## Our convention
Instructions to setup are below, but the code is generalized meaning you can feel free to customize it. We had set columns for variable info, one initial sourcing email draft, one followup draft, and added columns for keeping track of bounces and replies.  Our convention was to title sourcing email subjects "Berkeley Consulting Club Seeks {companyName} as a Client!" and followup drafts "Followup Template FA21", as reflected in the first few lines of code, but this is easily changeable and explained in the first comment. You can create multiple sourcing drafts and add a column to choose from which, structure titles differently, add additional tracking columns, etc. - just read the code!

## Initial setup
Most of the settings are generalized and can be changed easily at the top of the code.gs file. At a minimum, you must change the sheet ID. 
Make a copy of the email automation template sheet
Delete all the existing filled out rows those are just placeholders
Have to edit 2 sections of your sheet code
Click Tools -> Script Editor
Line 27, replace the ID in quotes with YOUR unique sheet id - this can be found in the url of your spreadsheet in between /d/ and /edit
Line 228, do the same thing
Create 2 drafts in your gmail inbox (same one that you are logged into when using sheets) based on the Sourcing Email Template and Followup Email Template, with the exact subject that’s in each respective templates
Make sure the word “here” in the sourcing email has the link to the client handbook attached in your drafts
Click on Tools -> Script Editor, then click on the alarm clock on the left, then click Add Trigger (bottom right), then fill out accordingly (this refreshes “days since reply”)


## Sending new emails
Populate each variable in the next available row
Ensure “email sent” cell is empty for rows you want to send, and “FollowUp?” = No
Mail Merge -> Send Emails -> enter the last row that you want the program to consider (or schedule send)
Don’t leave rows half-completed, or if you do, populate the “email sent” box with the keyword “skip” until complete so that the program doesn’t attempt to send an incomplete email; that would look very very bad

## Sending Follow-ups
Regularly check this document throughout the process, every hour the “Days Since Reply” are updated as well as the “Email Sent” column
The “Days Since Reply” column marks cells red if it has been > 3 days (only if the row isn’t a followup email, since this is supposed to be an indicator that you should followup)
For every row with a red “Date of Last Contact” cell, delete the date under the “email sent” column and change “FollowUp?” to = Yes
Now, click Mail Merge -> Send Emails and after successful sending, change the color back to clear (or schedule send)

## Schedule Send
Under the “For Schedule Send/Last Row” column (J2), enter the last row used (same input as what you enter when prompted while sending emails immediately)
Sending through the Mail Merge function will send automatically, if you’d like to schedule send a batch in the morning then click on Tools -> Script Editor, then click on the alarm clock on the left (“Triggers”), then click + Add Trigger (on the bottom right)
Fill it out accordingly (but instead of 8 you can schedule when you’d like military time), then click save



![Console Example](/pics/console.png?raw=true)
![Our Sourcing Draft](/pics/sourcing.png?raw=true)
![Our Followup Draft](/pics/followup.png?raw=true)