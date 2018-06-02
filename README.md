# email_utility
Generic email utility to do various things from the command line: mass emails, bulk imports/exports, etc.

This mostly exists because I needed to send bulk emails to people and I didn't like the existing
solutions available, and because many of the ones that would work aren't free. This is a simple
solution for my needs. Any extra functionality beyond bulk mailing is only there because I got bored
and figured why the hell not.

Functionality should be fairly obvious to anyone used to *nix command line tools. Possible exceptions:

* The spreadsheet of names--and this is because of my initial needs--has to have a specific format.
The header is the first row, and must be: Last_Name, First_Name, Middle_Initial, Email_Address, Phone, City, State,
Mail_Code. The next rows (with no blank rows) are filled in with the appropriate data. Again, this was because I
wrote this code to deal with a specific case. It would be easy enough to clean up, make generic, and use some
method of searching and pulling the column names to add the appropriate fields to the code. I just don't care enough
to do it.

I think everything else is mostly self-explanatory. I apologize in advance if it's not.