# Template-Email-Sender

It's a desktop application that allows you to send multiples emails for differents mailboxes, replacing some fields.

## Gmail mailbox configuration

Due to the application uses STMP to send the emails, you will need to [create an application code of your Gmail mailbox](https://support.google.com/mail/answer/185833?hl=en).

## Application inputs

- Excel file with column values for the fields to be replaced and the mailbox to send the email.
Every row is a different mailbox to send an email.

- HTML file with the mail structure structure where the fields to be replaced must be wrote as follow: *[ColumnNameInExcel].

