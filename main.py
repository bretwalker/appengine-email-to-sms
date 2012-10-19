import webapp2

from google.appengine.ext import db

from settings import APP_EMAIL

class Recipient(db.Expando):
    email = db.StringProperty()
    phone_number = db.StringProperty()
    last_message = db.StringProperty()
    
class MainHandler(webapp2.RequestHandler):
    def get(self):
        self.response.write('Who needs GOOD?')

class OutlookInstructions(webapp2.RequestHandler):
    def get(self):
        self.response.write("""<pre>
        ' Sends email containing reminder highlights
        ' unless reminder is in category "No Email Reminder"
        ' Edit email address to set to where email is sent
        Private Sub Application_Reminder(ByVal Item As Object)

          ' Get out of office status
          Dim bIsOOO As Boolean
          bIsOOO = False
          Dim oNS As Outlook.NameSpace
          Dim oStores As Outlook.Stores
          Dim oStr As Outlook.Store
          Dim oPrp As Outlook.PropertyAccessor

          Set oNS = Application.GetNamespace("MAPI")
          Set oStores = oNS.Stores
          For Each oStr In oStores
            If oStr.ExchangeStoreType = olPrimaryExchangeMailbox Then
              Set oPrp = oStr.PropertyAccessor
              bIsOOO = oPrp.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x661D000B")
            End If
          Next

          ' Do not send if out of office
          If bIsOOO = False Then
            ' Do not send if marked "No Email Reminder"
            If Item.Categories <> "No Email Reminder" Then
              Dim objMsg As MailItem
              Set objMsg = Application.CreateItem(olMailItem)

              objMsg.To = "%s"
              objMsg.Subject = "Reminder: " & Item.Subject
              objMsg.BodyFormat = olFormatPlain

              ' Code to handle the 4 types of items that can generate reminders
              Select Case Item.Class
                 Case olAppointment '26
                    If DateDiff("d", Item.Start, Item.End) = 0 Then
                        objMsg.Body = _
                          "Reminder: " & Item.Subject & vbCrLf & _
                          DatePart("m", Item.Start) & "/" & DatePart("d", Item.Start) & " -- " & _
                          Format(Item.Start, "h:mm ampm") & " - " & Format(Item.End, "h:mm ampm") & vbCrLf & _
                          "Location: " & Item.Location & vbCrLf
                    Else
                        objMsg.Body = _
                          "Reminder: " & Item.Subject & vbCrLf & _
                          "Start: " & Format(Item.Start, "h:mm ampm mm/dd/yyyy") & vbCrLf & _
                          "End: " & Format(Item.End, "h:mm ampm mm/dd/yyyy") & vbCrLf & _
                          "Location: " & Item.Location & vbCrLf
                    End If
                 Case olContact '40
                    objMsg.Body = _
                      "Reminder: " & Item.Subject & vbCrLf & _
                      "Contact: " & Item.FullName & vbCrLf & _
                      "Phone: " & Item.BusinessTelephoneNumber & vbCrLf & _
                      "Contact Details: " & vbCrLf & Item.Body
                  Case olMail '43
                    objMsg.Body = _
                      "Reminder: " & Item.Subject & vbCrLf & _
                      "Due: " & Item.FlagDueBy & vbCrLf
                  Case olTask '48
                    objMsg.Body = _
                      "Reminder: " & Item.Subject & vbCrLf & _
                      "Start: " & Item.StartDate & vbCrLf & _
                      "End: " & Item.DueDate & vbCrLf
              End Select

              objMsg.Send
              Set objMsg = Nothing
            End If
          End If
        End Sub
</pre>
        """ % APP_EMAIL)
        
app = webapp2.WSGIApplication([
    ('/', MainHandler),
    ('/outlook-code', OutlookInstructions)
], debug=False)
