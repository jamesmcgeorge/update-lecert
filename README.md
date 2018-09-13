# update-lecert

Powershell script designed to run as a scheduled task to update a Let's Encrypt certificate on an RD Gateway server

Uses ACMESharp https://github.com/ebekker/ACMESharp
Initial script cloned from https://marc.durdin.net/2016/11/automating-certificate-renewal-with-lets-encrypt-and-acmesharp-on-windows/

The server this script runs on is already updated with an ACME vault for the system account
