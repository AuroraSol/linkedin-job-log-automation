# <img width="29" height="30" alt="image" src="https://github.com/user-attachments/assets/0dbac7d4-8b6b-4ee7-bc1a-808407a49610" />  <img width="111" height="33" alt="image" src="https://github.com/user-attachments/assets/5619beb4-50e7-43a5-a80b-7e4b15b392e9" /><img width="23" height="33" alt="image" src="https://github.com/user-attachments/assets/ddb94ae6-d149-457a-b38e-81d016568c6c" /><img width="37" height="32" alt="image" src="https://github.com/user-attachments/assets/96fb6d1c-a71b-45b8-9a73-f6e2fabfd871" />



## Job Search Log using LinkedIn, Gmail and Google Sheets

A Google Apps Script that automatically parses LinkedIn job application emails and logs them into a Google Sheet with status tracking.

## Features
- **Auto-Log:** Automatically detects "Application Sent" emails and adds them to your Google Sheet.
- **Smart Updates:** If you receive a rejection email (e.g., "Your application to..."), the script finds the existing entry in your log and updates the status to `Declined - No Interview`.
- **Duplicate Prevention:** Marks processed emails as "Read" to ensure each application is only handled once.
---

### The Why?

I built this to keep an ongoing log of my job applications for better visibility and accountability. After hours of scanning descriptions and clicking "Apply", there were times I didn't actually finish an application maybe because of a schedule conflict or a sudden change in interest.

After a long day, it might feel like I applied for 20 roles, but in reality, it was only 12 because I exited halfway through. I created this script to cut out the manual work of copying and pasting job details into my log, ensuring my data is 100% accurate. Plus, it was a fun way to dive into script writing!

-----

## How to use

### What you'll need:

* Google sheets - Create a Job Search Log according to your needs.
  - Not sure what to add? I've added a free template ([here](https://docs.google.com/spreadsheets/d/188jRgjcdq90ZuZ1AfBR-D4e5KEfw8995HHjPW9S21Fw/edit?gid=0#gid=0))
* LinkedIn Account
* Gmail account - needs to be the email on your linked in, in order for you to recieve the application updates here

-----
### Steps to Set it Up 


* Inside of your Googlesheet, Click 'Extention' and Select Apps Script
  <img width="709" height="239" alt="image" src="https://github.com/user-attachments/assets/b23bd258-5f55-4b76-806a-034880634456" />




* Click on the Editor and '+' to create a new file, name it and then paste my code in
  <img width="654" height="205" alt="image" src="https://github.com/user-attachments/assets/2d3ce003-a35d-430e-a3dc-7cddf8a6d47d" />

* Click Save then Run
  <img width="446" height="57" alt="image" src="https://github.com/user-attachments/assets/9dbc39f3-1d3f-456c-9279-2de666c945a8" />

---
## Things to keep in mind

- **Unread Emails Only:** The script only looks at unread emails. If you want to backlog past applications, mark those emails as unread in Gmail first.
- **Column Mapping:** The script is set up for a specific column order:
  - Column A: Date
  - Column B: Method (LinkedIn)
  - Column C: Company Name
  - Column D: Job Title
  - Column F: Status
- **Triggers:** To fully automate this, set up a **Time-driven Trigger** in the Apps Script editor to run the `syncLinkedInApplications` function every hour.

---
### How to Create a Trigger - Running it based on time 

* In the Apps Script screen select "Triggers"

  <img width="278" height="314" alt="image" src="https://github.com/user-attachments/assets/668027fc-0ab1-4a80-968f-e3819d4d7b72" />

* Add Trigger
  <img width="1369" height="824" alt="image" src="https://github.com/user-attachments/assets/54ec8772-a7a3-45e5-8b84-6520eef7476f" />

* Select Event Source - Time Driven (make sure you have the correct function syncLinkedInApplications selected) 

  
  <img width="718" height="684" alt="image" src="https://github.com/user-attachments/assets/102329c6-c418-4c9a-b60a-e2f52cda0a3f" />

* Select Minutes or Hours (Whatever you prefer)

  
  <img width="692" height="794" alt="image" src="https://github.com/user-attachments/assets/eeb2ed0d-ac5d-4464-bbbc-c925d2a0d09e" />


* Select the interval you'd prefer, then Save

  
  <img width="706" height="834" alt="image" src="https://github.com/user-attachments/assets/89a99c1c-deda-4bff-9b35-c32d2687fccd" />



Give it a try and let me know what you think! I'm learning as I go so I'm always open to feedback. 






