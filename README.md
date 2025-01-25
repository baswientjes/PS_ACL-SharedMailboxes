# PS_ACL-SharedMailboxes
PowerShell script for dynamically assigning users to Shared Mailboxes in Microsoft 365, using an Excel workbook as input in the form of an ACL-like matrix.

In the Excel workbook, there is a hidden worksheet called #CONFIG# that contains two lists of user/group properties that can be used in the other worksheets.

In the top row you would put the email addresses of the shared mailboxes that you'd like the script to handle.

In row 2 and onwards, you can make user/group selections by filling in the cells in columns A-E.

Assign those selections to the shared mailboxes by putting R, S, or RS in the cells that cross a selection and a shared mailbox, much like an ACL matrix.
.
