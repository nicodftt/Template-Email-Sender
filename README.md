# Template-Email-Sender

It's a desktop application that allows you to send multiples emails for differents mailboxes replacing some fields.

## Gmail mailbox configuration

Due to the application uses STMP to send the emails, you will need to [create an application code of your Gmail mailbox](https://support.google.com/mail/answer/185833?hl=en).

## Application inputs

- Excel file with column values for the fields to be replaced and the mailbox to send the email.
Every row is a different mailbox to send an email.

- HTML file with the mail structure where the fields to be replaced must be wrote as follow: *[ColumnNameInExcel]

## Remaining implementations

- [X] Add button to set folder path for input files.

- [ ] Add functionality in order to allow  HTML files with images embedded.

- [ ] Improve user interface.

- [X] Add imputs for SMTP server and port.

- [ ] Add functionality and interface to edit text with template variables highlighted.

