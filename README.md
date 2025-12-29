Background

CHMeetings is widely used by churches around the world. It is a membership
management system that supports attendance tracking, event enrolment, and
financial management.

CHMeetings provides a fully featured API, which enables member information
to be retrieved and updated programmatically.

In the UK, churches are legally required to comply with safeguarding
policies. CHMeetings allows churches to configure custom fields to store
the expiry dates of DBS checks and safeguarding certificates.

This script integrates with the CHMeetings API to programmatically retrieve
member records and identify individuals whose DBS and/or safeguarding
certifications have expired or are due to expire within one month.

It generates an Excel report containing the relevant expiry details for
review by the People Team Coordinator and the Safeguarding Coordinator.

⸻

Prerequisites

CHMeetings must be configured with two custom fields to store the expiry
dates of DBS checks and safeguarding certificates.

API access to CHMeetings is required for the script to function.

The script uses the Google Gmail API to send email, rather than Google App
Passwords, which are obsolete, less secure, and increasingly deprecated
compared to modern OAuth-based authentication.

Users must create a project in Google Cloud Platform (GCP), enable the Gmail
API, configure the OAuth consent screen, create an OAuth client ID, and
generate a refresh token by following the official documentation:

https://developers.google.com/gmail/api/quickstart/python

⸻

Deployment

The script can be deployed on GCP, AWS, or GitHub Actions and executed on a
scheduled basis to ensure continuous monitoring and timely notifications.

To maintain security, sensitive parameters must be passed to the script
securely. When using GitHub Actions, secrets should be configured according
to the following guide:

https://docs.github.com/en/actions/how-tos/write-workflows/choose-what-workflows-do/use-secrets
