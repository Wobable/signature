# Signature
A VBS script that prompts for details to create a html signature in the default outlook folder.

This came about because I wanted to get all my work colleagues using the same style of MS Outlook signature, but it was too painful to recode the HTML signature for each of them.

An example of the signature is enclosed.

The script asks the user for the following details:

- Name
- Pronouns
- Pronounciation Guide
- Job title
- Phone number
- Working Hours
- Valediction (Sign off)

The script editor may also need to alter:
- where the final HTM file is stored
- the banner image (400x200 pixels)
- the banner link
- company name
- company address
- website
- twitter link
- linkedin link
- the flexible working policy
- the Welsh translation
- the legal text

The script assembles a HTM signature, saves it in their user folder ("C:\Users\[USERNAME]\AppData\Roaming\Microsoft\Signatures\") and opens a copy of it in their web browser.
You may need to open the signature in Outlook, edit it (minorly) in order to create a Rich Text and Plain Text version.
