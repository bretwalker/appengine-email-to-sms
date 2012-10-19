import hashlib
import logging
import re
import webapp2

from main import Recipient

from google.appengine.api import mail
from google.appengine.ext.webapp.mail_handlers import InboundMailHandler
from twilio.rest import TwilioRestClient

from settings import TWILIO_ACCOUT, TWILIO_TOKEN, TWILIO_NUMBER, FOOTER_STUFF1, FOOTER_STUFF2, FOOTER_STUFF3, AUTHORIZED_DOMAIN, ADMIN_EMAIL, APP_BASE_URL

class MailHander(InboundMailHandler):
    def receive(self, mail_message):
        sender = mail_message.sender
        
        logging.info("Received a message from: " + sender)
        
        if not sender.endswith(AUTHORIZED_DOMAIN) and not sender.endswith(AUTHORIZED_DOMAIN + '>') and not sender.endswith(ADMIN_EMAIL + '>') and sender != ADMIN_EMAIL:
            logging.info('Unauthorized domain')
            return
    
        r = Recipient.all().filter('email =', sender).get()
        
        if r is None:
            r = Recipient(email=sender)
            r.put()
            
            f = open('response_email.txt')
            response = f.read()
            
            mail.send_mail(sender="TXT Meeting Reminders<hi@txt-meeting.appspotmail.com>",
                          to=sender,
                          subject="Meeting reminders via text message",
                          body=response % APP_BASE_URL)
                          
            logging.info('Created user and sent instructions')
        elif hasattr(mail_message, 'subject') and re.match('^\d{10}$', re.sub("\D", "", mail_message.subject)):
            r.phone_number = re.sub("\D", "", mail_message.subject)
            r.put()
            
            self.send_sms('+1' + re.sub("\D", "", mail_message.subject), "I'll send you messages for " + r.email + '.  Send me an email with a phone number in the subject to change your number.') 
            logging.info('Added/updated phone number')
            
        elif hasattr(mail_message, 'subject') and re.sub("\W", "", mail_message.subject.lower()) == 'stop':
            r.delete()
            
            mail.send_mail(sender="TXT Meeting Reminders<hi@txt-meeting.appspotmail.com>",
                          to=sender,
                          subject="Meeting reminders via text message",
                          body="You've been unsubscribed and will no longer receive alerts via text message")
            logging.info('Deleted user')
        elif r.phone_number:            
            plaintext_bodies = mail_message.bodies('text/plain')
            plaintext_body = list(plaintext_bodies)[0][1].decode()
            
            m = hashlib.md5()
            m.update(plaintext_body.encode('utf-8'))
            hex_digest = m.hexdigest()
            
            if r.last_message <> hex_digest:
                r.last_message = hex_digest
                r.put()
                
                self.send_sms('+1' + r.phone_number, plaintext_body.replace(FOOTER_STUFF1, '').replace(FOOTER_STUFF2, '').replace(FOOTER_STUFF3, '').strip())

                logging.info('Sent SMS reminder')
            else:
                logging.info('Ignored duplicate reminder')
        else:
            mail.send_mail(sender="TXT Meeting Reminders<hi@txt-meeting.appspotmail.com>",
                          to=sender,
                          subject="Meeting reminders via text message",
                          body="I don't know where to send your reminders. Could you reply to this message with your phone number in the subject line?")
            logging.info('Sent email asking for phone number')
                
    def split_count(self, s, count):
        """Split string s at count, preserving words, returning list of strings.
        
        TODO: Handle long words/strings without spaces (not very relevant to this app)."""        
        if s <= count:
            return [s]
        
        strings = []
        current_string_length = 0
        current_string = []
        for word in s.split():
            current_string_length += len(word) + 1
            
            if current_string_length <= count:
                current_string.append(word)
            
            else:
                strings.append(' '.join(current_string))
                current_string_length = len(word) + 1
                current_string = []
                current_string.append(word)
                
        if len(current_string) > 0:
            strings.append(' '.join(current_string))

        return strings
                
    def send_sms(self, to, body):
        """Sends an SMS, spliting it into 160-character messages."""
        split_body = self.split_count(body, 155)
        client = TwilioRestClient(TWILIO_ACCOUT, 
                                  TWILIO_TOKEN)
                        
        if len(split_body) == 1:
            client.sms.messages.create(to=to,
                                       from_=TWILIO_NUMBER,
                                       body=split_body[0])
        else:
            i = 1
            for t in split_body:
                client.sms.messages.create(to=to,
                                           from_=TWILIO_NUMBER,
                                           body=t + ' (' + str(i) + '/' + str(len(split_body)) + ')')
                i += 1
                                           
app = webapp2.WSGIApplication([MailHander.mapping()], debug=True)