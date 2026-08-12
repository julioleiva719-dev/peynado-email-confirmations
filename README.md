# Appointment Confirmation Email System

Automated email confirmation system built with Google Apps Script for a law firm's appointment scheduling workflow.

# What it does

This system automatically sends branded, professional confirmation emails to clients when a new appointment email is entered into a Google Sheet — no manual work required.

- Triggers automatically when an email address is entered in the tracking sheet
- Pulls appointment details (client name, date, time, consultation type, attorney) directly from the row
- Sends a fully branded HTML email with the firm's logo, colors, and social media links
- Marks each row as "Sent" to prevent duplicate emails
- Includes a batch function to catch up on any pending confirmations

## Built with

- **Google Apps Script** (JavaScript-based automation for Google Workspace)
- **Google Sheets** as the data source and trigger
- **MailApp / HTML email templates** for branded, responsive email delivery

## How it works

1. An appointment is logged in the Google Sheet
2. When the client's email is entered in the designated column, an `onEdit` trigger fires
3. The script reads the relevant row, formats the date/time, and builds a personalized HTML email
4. The email is sent automatically and the row is marked as "Sent"
5. A separate function allows batch-sending confirmations for any pending rows

## 👤 Author

Built by Julio Leiva — Operations & Systems Automation
