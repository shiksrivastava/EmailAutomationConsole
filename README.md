# EmailAutomationConsole
Google Scripts Code to automate client acquisition sourcing emails, follow-ups, and check status

## Features
- send mass emails with variables in body per specified draft
- add variables in subject as well
- send followups in thread per another specified draft
- reports whether email was invalid (bounced) or sent
- automatically checks and updates the status if the email has been replied to

## Our convention
Instructions to setup are below, but the code is generalized meaning you can feel free to customize it. Some of our conventions that you can easily change in the first few lines:
- one initial sourcing email, and one followup email template
- columns for recipient name, days since reply, etc. 
- sourcing draft subject: "Berkeley Consulting Club Seeks {companyName} as a Client!"
- followup draft subject: "Followup Template FA21"
- last row that schedule send should check stored in Cell J2


## Initial setup
1) Make a new sheet and copy and paste code.gs to the script editor
2) change sheetID variable to your sheet ID (found in URL)
3) edit sheet columns to your necessary variables (can start with ours as an example)
![Console Example](/pics/console.png?raw=true)
4) Create a draft for your sourcing email and followup email (or additional), and change lines 17-21 to reflect the draft subjects. Here is our sourcing draft for reference:
![Our Sourcing Draft](/pics/sourcing.png?raw=true)
5) Add a script trigger (Tools -> Script Editor -> Alarm) according to refresh "days since reply" accordingly
![Console Example](/pics/refreshTrigger.png?raw=true)


## Sending new emails
1) Each row is an email - populate each variable for all the rows you want to send. 
2) Ensure “email sent” cell is empty for rows you want to send, and “FollowUp?” = No. If you would like to skip a row, enter the keyword "skip" in the "email sent" cell so that the program doesn't attempt to send an incomplete email
3) Mail Merge -> Send Emails -> enter the last row that you want the program to consider (or schedule send)

## Schedule Send
1) Under the “For Schedule Send/Last Row” column (J2), enter the last row used (same input as what you enter when prompted while sending emails immediately)
2) Add a trigger accordingly - change the time to your desired time of sending (in military time)
![Our Sourcing Draft](/pics/sendTrigger.png?raw=true)
