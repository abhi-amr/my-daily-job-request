# my-daily-job-request
Daily Cron to Send Emails with GitHub Actions &amp; Gmail

This repo :
1. Read rows from excel file
2. Extract Name, Email and Company
3. Format a msg using the above extracted details requesting for a job opportunity
4. Send the mail via smtp
5. This is daily cron jobs that would send only limited mails and bookmark the row of the last run to start in the next day
6. Logs are stored on log folder for each run
