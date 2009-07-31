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
import re
import pooler
import smtplib
import mimetypes
import binascii
import base64
import time
import os
from optparse import OptionParser
from email import Encoders
from email.Message import Message
from email.MIMEBase import MIMEBase
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
import netsvc
import sys
import poplib
import imaplib
import string
import email



if sys.version[0:3] > '2.4':
    from hashlib import md5
else:
    from md5 import md5
    

class poweremail_mailbox(osv.osv):
    _name="poweremail.mailbox"
    _description = 'Power Email Mailbox included all type inbox,outbox,junk..'
    _rec_name="subject"
    
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
                                ('followup','Follow Up')
                                ],'Folder'),
        'status':fields.selection([
                                ('read','Read'),
                                ('unread','Un-Read')
                                ],'Status'),
        'date_mail':fields.datetime('Rec/Sent Date')
    }

    _defaults = {

    } 
    
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
        if 'server_ref' in context.keys():
            #Get mails from that ID only 
            core_obj.get_fullmail(cr,uid,context['email_account'],context)
        else:
            raise osv.osv_except(_("Mail fetch exception"),_("No information on which mail should be fetched fully"))
                        
poweremail_mailbox()

    

