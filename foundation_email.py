#modules that need to be imported.
import string, csv, smtplib

#create variable data_reader and open csv file specifying excel (standard)
data_reader = csv.reader (open('fndver2.csv', 'rU'), dialect="excel")

# creates a new empty list variable
data = []

for row in data_reader:
    # add the row as a new entry in the list variable
    data.append(row)

# get the row of data from the list
headers = data.pop(0)

# creates a new, empty dictionary called i
i = {}


# fill the dictionary using each field of the header row
# set the value of each dictionary item to equal the header's index in the headers list
# this makes it possible to use a header name to call it's position b/c it's indexed.
for header in headers:
    i[header] = headers.index(header)


# this function takes a string, converts it into a number,
# then converts the number into currency (string)
# number before 'f' indicates decimal places, so in this case 0
def commas(string):
    if string is '':
        return '(information missing)'

    number = int(string)
    return '${:,.0f}'.format(number)

# formatting for numbers of grants approved and paid
# diff. than commas b/c not currency
def numbers(string):
    if string is '':
        return '(information missing)'

    number = int(string)
    return number
#this is a testing mechanism and prints specified individual email outputs w/o sending emails.
#this function performs a join on all strings in loop.
# whole_message = join of strings
# output = writes output to a text file in the files directory
# this function is for testing the output of an individual message in a txt format.
# def print_it(list_of_strings):
#    whole_message = ''.join(list_of_strings)
#    output = open('files/text.txt', 'w')
#    output.write(whole_message.encode('utf8'))
#    output.close()

#this is the log-in function; the server and hostname specifications are defined at end of script in send_mail
def log_in_to_mail_server(hostname, port, username, password, ssl=True):
    server = smtplib.SMTP(hostname, port)
    if ssl:
       server.starttls()
    server.login(username, password)

    return server

#this function passes arguments that structure the format of the email
def send_mail(server, from_address, to_address, subject, body):
    return server.sendmail(
        from_address,
        to_address,
        ('From: {from_address}\r\nTo: {to_address}\r\nSubject: {subject}\r\n\r\n'
            '{body}\r\n').format(
                from_address=from_address, to_address=to_address,
                subject=subject, body=body))

#all_emails creates a list of the text [], which contains the structure/format of one email. 
#all_emails therefore is a list that contains all the emails.
all_emails = []
#all_email_address is a list
all_email_addresses = []
for row in data:

#this appends the row containing email addresses [3] in the csv file to the all_email_addresses list
    all_email_addresses.append(row[3])

#this is the list that acts as the template for the individual email messages.
    text = []
    text.append('%s\n' % row[5]) #org_name
    text.append('March 11, 2014\n\n')
    text.append('%s:\n\n' % row[1]) #name of person
    text.append("""Thank you for taking the time to participate in our report on grant maker practices for our 
annual report on Foundations. You provided us with the following information. Please verify that this 
is correct.\n\nYou can send updates, including filling in missing information, by replying to this email. 
Even if no changes are required, please reply to the email letting us know. Please reply no later than COB,
Wednesday, March 12th.\n\n""")
    text.append('Please note, all information for FY13 and FY14 will be presented as estimates to readers.\n\n\n')

    # since we have to determine if the 990 exists for everything,
    # do that once here and save the result as a Boolean
    if row[i['asset12']] is '0':
        has990 = False
        text.append('We do not have your FY-12 990 on file. Please direct us to where we can access a copy.')
        text.append('\n')
        text.append('\n')
    else:
        has990 = True

    ### ASSETS
    # put the initial string at the end of our list
    text.append('Total assets (fair market value) at end of fiscal year (equivalent to line I on the Form 990-PF):')
    text.append('\n')
    #this is the first instance of 'u'\u2022' which is unicode for a bullet point.
    #see at end that joins have to encode on utf-8 for python to read this properly.
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[7]))

    if row[i['nomatch_asset']] is not '':
        text.append('\n')
        text.append('The FY12 number you reported does not match what\'s in your 990 (%s)' % commas(row[9]))

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[10]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[11]))
    text.append('\n')
    text.append('\n')


    ### ADMIN COSTS
    text.append('Total operating and administrative costs (equivalent to Part I, Line 24d on the Form 990-PF):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[16]))

    if row[i['nomatch_admin']] is not '':
        text.append('\n')
        text.append('The FY12 number you reported does not match what\'s in your 990 (%s)' % commas(row[17]))

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[18]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[19]))
    text.append('\n')
    text.append('\n')

    ### COMPENSATION COSTS
    text.append('Compensation of officers, directors, trustees, etc. (equivalent to Part I, Line 13d on the Form 990-PF):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[20]))

    if row[i['nomatch_comp']] is not '':
        text.append('\n')
        text.append('The FY12 number you reported does not match what\'s in your 990 (%s)' % commas(row[21]))

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[22]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[23]))
    text.append('\n')
    text.append('\n')

    ### GRANTS PAID
    text.append('Total value of grants paid (equivalent to Part I, line 25d on the Form 990-PF):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[24]))

    if row[i['nomatch_grntpd']] is not '':
        text.append('\n')
        text.append('The FY12 number you reported does not match what\'s in your 990 (%s)' % commas(row[25]))

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[26]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[27]))
    text.append('\n')
    text.append('\n')

    ### PRI
    text.append('Total value of program-related investments (enter 0 if nothing was paid, equivalent to Part XII, line 1b on the Form 990-PF):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[28]))

    if row[i['nomatch_pri']] is not '':
        text.append('\n')
        text.append('The FY12 number you reported does not match what\'s in your 990 (%s)' % commas(row[29]))

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[30]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[31]))
    text.append('\n')
    text.append('\n')

    ### GRANTS APPROVED
    text.append('Total value of grants approved (including program-related investments):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % commas(row[32]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % commas(row[33]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % commas(row[34]))
    text.append('\n')
    text.append('\n')

    ### NUMBER GRANTS APPROVED
    text.append('Total number of grants approved (including program-related investments):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % numbers(row[35]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % numbers(row[36]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % numbers(row[37]))
    text.append('\n')
    text.append('\n')

    ### NUMBER GRANTS PAID
    text.append('Total number of grants paid (including program-related investments):')
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY12 reported: %s' % numbers(row[38]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY13 reported: %s' % numbers(row[39]))
    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('FY14 reported: %s' % numbers(row[40]))
    text.append('\n')
    text.append('\n')

    ### EST 2014 GIVING
    text.append('Estimated 2014 Giving. Choices included: Increase by 3 percent or more, Decrease by 3 percent or more, Remain the same, Don\'t know:')
    
    #this is the last point for verification. not a number or currency.
    #did not set a function; instead established an if statement.
    last_msg = row[41]

    if last_msg is '':
        # putting a default selection, in case they didn't select anything
        last_msg = 'No option selected'

    text.append('\n')
    text.append('\t'u'\u2022')
    text.append('Response: %s' % last_msg)
    text.append('\n')
    text.append('\n')
    text.append('\n')

    ### SIGNATURE
    text.append('Thank you for taking the time to review and verify this information. Please respond no later than March 12th.')
    text.append('\n')
    text.append('\n')
    text.append('Sarah Frostenson')
    text.append('\n')
    text.append('The Chronicle of Philanthropy')
    text.append('\n')
    text.append('sarah.frostenson@philanthropy.com')
    text.append('\n')
    text.append('202-466-1201')

    # end of loop, add text list to all_email list
    # this is so we can save all emails and access them later to send an email.
    all_emails.append(text)

#this is for the print_it function for testing purposes.
#print_it(all_emails[1])

#this is the specifications for our function log_in_to_mail_server. plug in appropriate variables.
server = log_in_to_mail_server('HOSTNAME GOES HERE', 'PORT GOES HERE', 'USERNAME GOES HERE', 'PASSWORD GOES HERE') 

#this uses the zip function to match the list all_emails with the corresponding all_email_addresses. 
#zip works on matching the content of two lists. and creates email for all_emails and email_address for all_email_addresses.
#establishes a safety catch if emails fail to send with try and except
#there is also a series of debugging mechanisms here to determine errors if they arise.
for email, email_address in zip(all_emails, all_email_addresses):
    try:
        #this portion refers back to the send_mail function
        #it also bccs research with each of the emails sent in the email address list.
       send_mail(server, 'research@philanthropy.com', [email_address, 'research@philanthropy.com'], '2014 Chronicle of Philanthropy Foundation Report: Verify Key Data Points by March, 12th', ''.join(email).encode('utf-8'))
    except smtplib.SMTPRecipientsRefused:
        print 'Recipients refused when sending email to ' + email_address
    except smtplib.SMTPSenderRefused:
        print 'Sender refused when sending email to ' + email_address
    except smtplib.SMTPDataError:
        print 'Unknown SMTP data error when sending email to ' + email_address
    except smtplib.SMTPServerDisconnected:
        print 'Server disconnected when sending email to ' + email_address + '; retrying.'  
        server = log_in_to_mail_server('HOSTNAME GOES HERE', 'PORT GOES HERE', 'USERNAME GOES HERE', 'PASSWORD GOES HERE')
        server.connect()
        send_mail(server, 'research@philanthropy.com', [email_address, 'research@philanthropy.com'], '2014 Chronicle of Philanthropy Foundation Report: Verify Key Data Points by March, 12th', ''.join(email).encode('utf-8'))
    except Exception:
        print 'Other unknown error when sending email to ' + email_address
        import pdb; pdb.set_trace()

# TESTING ONLY--TAKE THIS OUT
#this sends one email to a specified email with research cc'ed to test formating, etc.
#send_mail(server, 'research@philanthropy.com', ['sarah.e.frostenson@gmail.com', 'research@philanthropy.com'], '2014 Chronicle of Philanthropy Foundation Report: Verify Key Data Points by March, 12th', ''.join(all_emails[3]).encode('utf8'))

