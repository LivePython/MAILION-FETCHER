# MAILION-FETCHER

This application is capable of extracting contacts in any email account. 
 
* This mailing fetcher is an application capable of fetching all contacts in an email address.
* It fetches all emails from every folder (inbox, draft, sent, even junk, trash, and user custom folder) of the email.
* Then the fetcher parses every email’s header and gets the email address of the sender or receiver (and Bcc and cc address if it exists) and saves them as result in the file(s).
* This application only takes the email address and password, port or SMTP server not needed at all, and logs into the account to extract all to, cc, and bcc contacts in the address. It can change emails like office365, OWA, etc (excluding Gmail and Yahoo).
