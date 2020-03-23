# cvpReport
Custom report for CloudVision Portal

# Description
This script connects to CloudVision Portal and then does the following:

1- Gets the hostname, model, SW Version, IP Address, Serial Number, Up Time

2- Calculates the Daily Availability, CPU load , Free Memory from the last 96 x 15 minute period values.

3- Gets the number of completed tasks per user in the last 7 days.

4- Reports all in an excel file and sends an email with the file in attachment.
