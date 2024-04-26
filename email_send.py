import win32com.client as win32
import os

# Read email addresses from a text file
file_path = r'emails.txt'  # Change to your file's path
with open(file_path, 'r') as file:
    emails = [line.strip() for line in file if line.strip()]

# Check if Outlook is installed
try:
    outlook = win32.Dispatch('outlook.application')
except Exception as e:
    print("Error: Outlook application could not be accessed.")
    print(e)
    exit()

# Compose the email
subject = "Don't Miss Out: No Contracts, Free Setup, & Lowest Rates in Canada!"

# HTML body with bold text
for email in emails:
    html = f"""\
        <html>
<body style="font-family: Arial, sans-serif; font-size: 15px;">
<p> Good day, </p>

<p> Hope you’re doing well. </p>

<p> I’m reaching out to you in regards to your merchant processing services. </p>
<p> Tired of high processing fees? Our limited-time offer is here to give your business a boost without draining your budget! </p>
<p> Here's what you can expect from us:</p>

<P style="color:red; margin-left: 10px;"> • <b> Unbeatable Rates: </b> Lifetime guaranteed fees starting as low as <b> 0.10% </b> for Visa and MasterCard. </p>
<P style="color:red; margin-left: 10px;"> • <b> Advanced Terminals: </b> Get the latest GoDaddy/POYNT Smart Terminal for just <b> $30/month </b>, or choose any other terminal of your choice. </p>
<P style="color:red; margin-left: 10px;"> • <b> Low-Cost Debit: </b> Just <b> $0.04 </b> per debit transaction. </p>
<P style="color:red; margin-left: 10px;"> • <b> No Setup Fees: </b> Try our services without any upfront costs—no strings attached. </p>
<P style="color:red; margin-left: 10px;"> • <b> No Contracts: </b> Cancel any time with <b> zero penalties </b>. You’re in control! </p>
<P style="color:red; margin-left: 10px;"> • <b> Quick Payouts: </b> Next-day funding to keep your cash flow smooth. </p>
<P style="color:red; margin-left: 10px;"> • <b> 24/7 Customer Support: </b> Get immediate help without endless call trees. </p>
<P style="color:red; margin-left: 10px;"> • <b> Price Protection: </b> We guarantee your rates will never increase. </p>
<P style="color:red; margin-left: 10px;"> • <b> Low Admin Fee: </b> Just $5/month to cover administrative costs. </p>

<p style="color: black; font-size: 20px;"> <b> Check Out Your Savings </b> </p>

<P> Simply attach your recent merchant statement, and we'll do a free side-by-side comparison to show how much you could save. </p>
<p> We’re excited to help you take your business to new heights with our reliable and affordable merchant processing services. Got questions? Just hit 'reply' or give us a call. </p>
<p> Sincerely, </p>
<p style="margin: 0.25px 0;"><b>Zack Nelson</b></p>
<p style="margin: 0.25px 0;">Sales Representative</p>
<p style="margin: 0.25px 0;">Email: <a href="mailto:zack@rarepayments.com">zack@rarepayments.com</a></p>
<p style="margin: 0.25px 0; color:red;"><b>Phone: +1 (807) 500 5520</b></p>
<p style="margin: 0.25px 0;"><b>Rare Payments MSP/ISO 2023</b></p>
          </body>
        </html>
        """

    # Send email to each address
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = subject
    mail.HTMLBody = html
    mail.Send()

print("Emails have been sent successfully!")
