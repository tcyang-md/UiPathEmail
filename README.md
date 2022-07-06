# UiPath Email :mailbox_with_mail:
Automatically composes a tailored-to-sender email using UiPath by reading in data from `EmailDb.xlsx`, editing `EmailTemplateEdit.docx` to personalize the email, exports as `EmailTemplateEdit.html`, and sends it after Gmail authentication.

## About
I created this project to help me send fundraising and thank you emails. In order to help our constituents feel more connected to our organization, we usually address the email by the first and last name "Dear firstname lastname," and sign off the email with our name, camp name, and our position in the organization (Chris Lamb Yang, Development Co-coordinator). This is an extremely tedious task as normally I have to copy each email address, change the name, then add a subject line, etc... My program thankfully does it all for you. 
Furthermore, we get donations ranging from a couple of dollars to a couple thousand dollars, so in order to thank each of our donors appropriately, we have different email templates that we use. For example, for anyone that donates under $100, we just write a simple thank you email, but for someone that donates over $100, we usually send them a personalized thank you notes and other goodies so we'd like to let them know exactly how we are going to setward them in the email. The program does this by reading in the `Donation Amount` column in the spreadsheet and selects the appropriate template and fills out all the personal information. 

## How it works
1. `Main` will ask you to sign in to Gmail through OAuth2
2. Read in an excel file `EmailDb.xlsx` row by row and obtain all the email data from the row (`firstname`, `lastname`, `email`, `sender`, `title`, `subject`, `donation`)
3. Then it will make a copy of the appropriate email template and save it as a copy (`EmailTemplate$amt.docx` -> `EmailTemplateEdit.docx`)
4. Change keywords in the `EmailTemplateEdit.docx` such as `{{ firstname }}` to the contents of the variable `firstname` maintaining original word document formatting
5. Export as `EmailTemplateEdit2.html` so that the send email function can attach it as the body of the email
6. Send the email!

## Instructions
- Download [UiPath Studio](https://account.uipath.com/login?state=hKFo2SBUUG5jSkFpMTliMVBXM2FEUTVUZ1d5XzJuTW9QTTJsN6FupWxvZ2luo3RpZNkgYVo3OVRlQ1lxalJYY242b0JHVlZOT3ctSE1aZ052bFGjY2lk2SAyeXQ5SGRGNDVPMDA2SDlxZFBjUDlhczVjZEdibkNXcw&client=2yt9HdF45O006H9qdPcP9as5cdGbnCWs&protocol=oauth2&audience=https%3A%2F%2Fuipath.eu.auth0.com%2Fapi%2Fv2%2F&scope=openid%20profile%20email%20read%3Acurrent_user%20update%3Acurrent_user_metadata&redirect_uri=https%3A%2F%2Fcloud.uipath.com%2Fportal_%2FauthCallback&type=signup&platform_name=UiPath%20Automation%20Cloud%20for%20enterprise&ecommerceRedirect=false&retryUrl=&product_name=UiPath%20Automation%20Cloud&company_code=B2B_CP&enable_marketing_fields=true&cloudrpa_signup_subdomain=%2Fportal_&register_endpoint=%2Fregister&use_local_registration=false&response_type=code&response_mode=query&nonce=cHBWNWx6ZHVWNjQxMDhGUUdxZ2pPeDhhMmpscFc0djNrQ2o1U1o0UkVscA%3D%3D&code_challenge=fH8t1Qj0_pTxEHKFcn-dvsrsmCd9xKH_5XSCANURoRY&code_challenge_method=S256&auth0Client=eyJuYW1lIjoiYXV0aDAtcmVhY3QiLCJ2ZXJzaW9uIjoiMS4yLjAifQ%3D%3D) and create an account
- Download and move all the repository's contents into your directory
- Go to the ribbon and click on "Manage Packages"
- Install these packages:
  - UiPath.Excel.Activities
  - UiPath.System.Activities
  - UiPath.Mail.Activities
  - UiPath.Word.Activities
  - UiPathTeam.Gmail.Activities
  - UiPath.UIAutomation.Activities
- Edit `EmalDb.xlsx` and `EmailTemplate.docx` documents, `EmailTemplateEdit.docx` and `EmailTemplateEdit.html` will automatically be generated and the examples will be overridden.
