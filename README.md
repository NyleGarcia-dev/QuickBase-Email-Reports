# QuickBase-Email-Reports
Enables Sending Quickbase reports via Gmails Imap Functionality 


Step 1:
Download and install Requierments

Step 2:
Edit and rename Config file.

remove -example from the end.

Step 3:
Run Reporting.py to generate Google auth

Example usage:

python3 reporting.py 'Subject name' <ReportID> <TableID> ExampleEMAIL1@Example.com Example@Example.com ......

python3 reporting.py 'potato Report' 22 123457 Example@Example.com 

Cron:

0 6 * * 1 python3 reporting.py 'potato Report' 22 123457 Example@Example.com 

