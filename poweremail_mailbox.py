#########################################################################
#Power Email is a module for Open ERP which enables it to send mails    #
#Core settings are stored here                                          #
#########################################################################
#   #####     #   #        # ####  ###     ###  #   #   ##  ###   #     #
#   #   #   #  #   #      #  #     #  #    #    # # #  #  #  #    #     #
#   ####    #   #   #    #   ###   ###     ###  #   #  #  #  #    #     #
#   #        # #    # # #    #     # #     #    #   #  ####  #    #     #
#   #         #     #  #     ####  #  #    ###  #   #  #  # ###   ####  #
# Copyright (C) 2009  Sharoon Thomas                                    #
#                                                                       #
#This program is free software: you can redistribute it and/or modify   #
#it under the terms of the GNU General Public License as published by   #
#the Free Software Foundation, either version 3 of the License, or      #
#(at your option) any later version.                                    #
#                                                                       #
#This program is distributed in the hope that it will be useful,        #
#but WITHOUT ANY WARRANTY; without even the implied warranty of         #
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          #
#GNU General Public License for more details.                           #
#                                                                       #
#You should have received a copy of the GNU General Public License      #
#along with this program.  If not, see <http://www.gnu.org/licenses/>.  #
#########################################################################

from osv import osv, fields
   

class poweremail_mailbox(osv.osv):
    _name="poweremail.mailbox"
    _description = 'Power Email Mailbox included all type inbox,outbox,junk..'
    _rec_name="subject"

    def get_headers(self,cr,uid,context={}):
        #email_account: THe ID of poweremil core account
        #Context should also have the last downloaded mail for an account
        #Normlly this function is expected to trigger from scheduler hence the value will not be there
        core_obj = self.pool.get('poweremail.core_accounts')
        if 'email_account' in context.keys():
            #Get mails from that ID only
            core_obj.get_headers(cr,uid,context['email_account'])
        else:
            accounts = core_obj.search(cr,uid,[('state','=','approved')])
            #Now get mails for each account
            for account in accounts:
                core_obj.get_headers(cr,uid,[account])

    def get_fullmail(self,cr,uid,context={}):
        core_obj = self.pool.get('poweremail.core_accounts')
        if 'mailboxref' in context.keys():
            #Get mails from that ID only
            core_obj.get_fullmail(cr,uid,context['email_account'],context)
        else:
            raise osv.osv_except(_("Mail fetch exception"),_("No information on which mail should be fetched fully"))
        
    def send_all_mail(self,cr,uid,ids,ctx={}):
        #get ids of mails in outbox
        ids = self.search(cr,uid,[('folder','=','outbox')])
        #send mails one by one
        for id in ids:
            core_obj=self.pool.get('poweremail.core_accounts')
            values =  self.read(cr,uid,ids[0],[])
            
            if core_obj.send_mail(cr,uid,ids,[values['pem_to']],[values['pem_cc']],[values['pem_bcc']],[values['pem_subject']],[values['pem_body_text']],body_html="",payload={})):
                self.write(cr,uid,id,{'folder':'sent'})

    _columns = {
            'pem_from':fields.char('From', size=64),
            'pem_to':fields.char('Recepient (To)', size=64),
            'pem_cc':fields.char(' CC', size=64),
            'pem_bcc':fields.char(' BCC', size=64),
            'pem_subject':fields.char(' Subject', size=200),
            'pem_body_text':fields.text('Standard Body (Text)'),
            'pem_body_html':fields.text('Body (Text-Web Client Only)'),
            'pem_attachments_ids':fields.many2many('ir.attachment', 'mail_attachments_rel', 'mail_id', 'att_id', 'Attachments'),
            'pem_account_id' :fields.many2one('poweremail.core_accounts', 'User account'),
            'server_ref':fields.integer('Server Reference of mail',help="Applicable for inward items only"),
            'pem_recd':fields.char('Received at',size=50),
            'mail_type':fields.selection([
                                    ('multipart/mixed','Has Attachments'),
                                    ('multipart/alternative','Plain Text & HTML with no attachments'),
                                    ('text/plain','Plain Text'),
                                    ('text/html','HTML Body'),
                                    ],'Mail Contents'),
            'folder':fields.selection([
                                    ('inbox','Inbox'),
                                    ('drafts','Drafts'),
                                    ('outbox','Outbox'),
                                    ('trash','Trash'),
                                    ('followup','Follow Up'),
                                    ('sent','Sent Items'),
                                    ],'Folder'),
            'status':fields.selection([
                                    ('read','Read'),
                                    ('unread','Un-Read')
                                    ],'Status'),
            'date_mail':fields.datetime('Rec/Sent Date')
        }

    _defaults = {

    } 



poweremail_mailbox()

    

