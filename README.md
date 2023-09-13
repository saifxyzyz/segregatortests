Made By: 

<a href="https://github.com/b1gh3ro" ><img src="https://github.com/Mystic-Arts/Automated_Segregator/assets/73358222/6d0e8929-9cd2-4ed3-aa56-859dc3797dc8" style="width:13vw" /></a>
<a href="https://github.com/Mohammed-Shoaib01" target="_blank"><img src="https://github.com/Mystic-Arts/Automated_Segregator/assets/73358222/b4c7fb3f-326d-4e93-b868-16feb45a36d0" style="width:13vw"/></a>
<a href="https://github.com/saifxyzyz" target="_blank"><img src="https://github.com/Mystic-Arts/Automated_Segregator/assets/73358222/28fec48c-7279-413e-a068-0f1c6f006135" style="width:13vw"/> </a>



@b1gh3ro, @Mohammed-Shoaib01, @saifxyzyz





# Steps to get started on a new environment

1. Download or clone repository
2. Run "pip install -r requirements.txt"
3. Create two folders named "input" and "files"



This is a bot made for amazon Po automation. To make the process easier with lesser manual effort.

it starts with the bot connecting to the email client. This is done by the IMAP4 protocol client (**imaplib2**) in python.

This module defines three classes, IMAP4, IMAP4_SSL and IMAP4_stream, which encapsulate a connection to an IMAP4 server.(This module defines three classes, IMAP4, IMAP4_SSL and IMAP4_stream, which encapsulate a connection to an IMAP4 server)

First it connects to the client using the **IMAP4_SSL** function and logs into the client.

from here it goes into idle mode using imap's **idle** function. Here it checks constantly for any **unread** email to be shown.

once it finds an unread email it divides it into multiple parts.

It extracts the subject of the email and stores it the **subject** variable.

it then checks for any attachments and starts looking in the **subject** variable for arguments.

**arguments** (to be provided in the email subject)

1.  segregate_file: this is given when you need the file to be divided into multiple different files according to the seller name
    Usage: **action:"segregate_file"**
    Inside the segregate_file argument you have some more sub arguments:**arguments** (to be provided in the email subject)

        i)  availibility: given when you only need to check with the seller for the availibility of products. This formats the files and sends them to sellers for checking
        **usage: type:"avail"**
        ii) Purchase Order: Given after the seller has replied with the availibility. This formats the files and sends them to the sellers to actually buy the given products.
        **usage: type:"purchase"
        iii)nothing: When nothing is given for type then it just segregates it

2.  Change_file: this adds a new file containing seller names and barcodes for segregation. The name of the new file should also be Book.xlsx
    **usage: action:"Change_file"**

3.Purchase Order Number: the purchase order number is to be given to be saved in that manner.
**Usage: PONo:"<YourPurchaseOrderNumber>"**

Here the main folders are created according to the PO provided.
According to the type, more folders are created for the **type:""** arguments inside the po folder.

**From here every argument is passed down to the bot to segregate the file**

**bot()**

The bot makes use of the **openpyxl** library. **openpyxl** is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

This has access to the "Book.xlsx" file with barcodes mapped to the seller name. and all the arguments we passed from above. Ibart makes a dict with seller names and the barcodes that they have.

It opens the file provided and has a list named findMainBar to find the barcode in the sheet. **As of now it only checks the first columns for barcode (should probably fix that) **

It checks the provided sheet and the seller name corrosponding to the seller name. Here it starts making new folders and files.

the files are named after the seller and they are placed in avail purchase or directly in the po no folder.

It copies the template and goes to append the the data.

after this it performs calculations.

it then changes the placeholders with real data.

It now calls the sendersf()

**sendersf():**

this uses the **smtplib** library. The smtplib module defines an SMTP client session object that can be used to send mail to any internet machine with an SMTP or ESMTP listener daemon.

along with **MIMEMultipart**. The MimeMultipart class is an implementation of the abstract Multipart class that uses MIME conventions for the multipart data.
Like before when we split the email into multiple parts here we rejoin them to send it.

it requires the subject, from ,to, and attachment.

It logs in throught the smtp client and sends the multple emails attached.

Now we come to the change file part:
**usage: action:"Change_file"**

**The given file contains the information to map the seller names to their respective barcodes**.

We realised that we would need to be able to update the main file for this project to be functional. As such, you are able to change the file by this command.

The email sent should have another mapping file to be attached to it. And it has to be named as **Book.xlsx** or else the user would be sent an error as response. If everything went as expected you would get a reply stating the same.
