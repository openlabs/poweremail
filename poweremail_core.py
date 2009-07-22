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

class poweremail_core_accounts(osv.osv):
    _name = "poweremail.core_accounts"

    _columns = {
        'name': fields.char('Email Account Desc', size=64, required=True, select=True),
        'user':fields.many2one('res.users','Related User',required=True,readonly=True),
        
        'email_id': fields.char('Email ID',size=120,required=True),
        
        'smtpserver': fields.char('Server', size=120, required=True),
        'smtpport': fields.integer('SMTP Port ', size=64, required=True),
        'smtpuname': fields.char('User Name', size=120, required=True),
        'smtppass': fields.char('Password', size=120, invisible=True, required=True),
        'smtpssl':fields.boolean('Use SSL'),
        
        'iserver':fields.char('Incoming Server',size=100),
        'isport': fields.integer('Port'),
        'isuser':fields.char('User Name',size=100),
        'ispass':fields.char('Password',size=100),
        'iserver_type': fields.selection([('imap','IMAP'),('pop3','POP3')], 'Server Type'),
        'isssl':fields.boolean('Use SSL'),
        'isfolder':fields.char('Folder',readonly=True,size=100,help="Folder to be used for downloading IMAP mails\nClick on adjacent button to select from a list of folders"),
        
        'company':fields.selection([
                                    ('yes','Yes'),
                                    ('no','No')
                                    ],'Company Mail A/c',help="Select if this mail account does not belong to specific user but the organisation as a whole. eg:info@somedomain.com",required=True,),

        'state':fields.selection([
                                  ('draft','Initiated'),
                                  ('awaiting','Awaiting Approval'),
                                  ('suspended','Suspended'),
                                  ('approved','Approved')
                                  ],'Account Status',required=True,readonly=True),
                }

    _defaults = {
         'name':lambda self,cr,uid,ctx:self.pool.get('res.users').read(cr,uid,uid,['name'])['name'],
         'smtpserver':lambda *a:'smtp.gmail.com',
         'smtpport':lambda *a:587,
         'smtpuname':lambda *a:'yourname@yourdomain.com',
         'email_id':lambda *a:'yourname@yourdomain.com',
         'smtpssl':lambda *a:True,
         'state':lambda *a:'draft',
         'user':lambda self,cr,uid,ctx:uid
                 }
                 
    _sql_constraints = [
        ('email_uniq', 'unique (email_id)', 'Another setting already exists with this email ID !')
    ]
    def _constraint_unique(self, cr, uid, ids):
        print self.read(cr,uid,ids,['company'])[0]['company']
        if self.read(cr,uid,ids,['company'])[0]['company']=='no':
            print self.read(cr,uid,ids,['email_id'])[0]['email_id']
            accounts = self.search(cr, uid,[('user','=',uid),('company','=','no')])
            if len(accounts) > 1 :
                return False
            else :
                return True
        else:
            return True
#Reinserted # Constraint removed. Think its meaningless
    _constraints = [
        (_constraint_unique, _('Error: You are not allowed to have more than 1 account.'), [])
    ]

    def on_change_emailid(self, cr, uid, ids, smtpacc=None,email_id=None, context=None):
        self.write(cr, uid, ids, {'state':'draft'}, context=context)
        return {'value': {'state': 'draft','username':email_id,'isuser':email_id}}

    def get_approval(self,cr,uid,ids,context={}):
        #TODO: Request for approval
        self.write(cr, uid, ids, {'state':'awaiting'}, context=context)
#        wf_service = netsvc.LocalService("workflow")

    def do_approval(self,cr,uid,ids,context={}):
        #TODO: Check if user has rights
        self.write(cr, uid, ids, {'state':'approved'}, context=context)
#        wf_service = netsvc.LocalService("workflow")

    def get_reapprove(self,cr,uid,ids,context={}):
        #TODO: Check if user has rights
        self.write(cr, uid, ids, {'state':'draft'}, context=context)

    def do_suspend(self,cr,uid,ids,context={}):
        #TODO: Check if user has rights
        self.write(cr, uid, ids, {'state':'suspended'}, context=context)    

poweremail_core_accounts()

class poweremail_core_selfolder(osv.osv_memory):
    _name="poweremail.core_selfolder"
    _description = "Shows a list of IMAP folders"

    def makereadable(self,imap_folder):
        if imap_folder:
            result = re.search(r'(?:\([^\)]*\)\s\")(.)(?:\"\s)(?:\")([^\"]*)(?:\")', imap_folder)
            seperator = result.groups()[0]
            folder_readable_name = ""
            splitname = result.groups()[1].split(seperator) #Not readable now
            if len(splitname)>1:#If a parent and child exists, format it as parent/child/grandchild
                for i in range(0,len(splitname)-1):
                    folder_readable_name=splitname[i]+'/'
                folder_readable_name = folder_readable_name+splitname[-1]
            else:
                folder_readable_name = result.groups()[1].split(seperator)[0]
            return folder_readable_name
        return False
    
    def _get_folders(self,cr,uid,ctx={}):
        print cr,uid,ctx
        if 'active_ids' in ctx.keys():
            record = self.pool.get('poweremail.core_accounts').browse(cr,uid,ctx['active_ids'][0])
            print record.email_id
            if record:
                folderlist = []
                try:
                    if record.isssl:
                        serv = imaplib.IMAP4_SSL(record.iserver,record.isport)
                    else:
                        serv = imaplib.IMAP4(record.iserver,record.isport)
                except imaplib.IMAP4.error,error:
                    raise osv.except_osv(_("IMAP Server Error"), _("An error occurred : %s ") % error)
                try:
                    serv.login(record.isuser, record.ispass)
                except imaplib.IMAP4.error,error:
                    raise osv.except_osv(_("IMAP Server Login Error"), _("An error occurred : %s ") % error)
                try:
                    for folders in serv.list()[1]:
                        folder_readable_name = self.makereadable(folders)
                        if folders.find('Noselect')==-1: #If it is a selectable folder
                            folderlist.append((folder_readable_name,folder_readable_name))
                        if folder_readable_name=='INBOX':
                            self.inboxvalue = folder_readable_name
                except imaplib.IMAP4.error,error:
                    raise osv.except_osv(_("IMAP Server Folder Error"), _("An error occurred : %s ") % error)
            else:
                folderlist=[('invalid','Invalid')]
        else:
            folderlist=[('invalid','Invalid')]
        return folderlist
    
    _columns = {
        'name':fields.many2one('poweremail.core_accounts',string='Email Account',readonly=True),
        'folder':fields.selection(_get_folders, string="IMAP Folder"),
        
    }
    _defaults = {
        'name':lambda self,cr,uid,ctx: ctx['active_ids'][0],
        'folder': lambda self,cr,uid,ctx:self.inboxvalue
    }

    def sel_folder(self,cr,uid,ids,context={}):
        if self.read(cr,uid,ids,['folder'])[0]['folder']:
            if not self.read(cr,uid,ids,['folder'])[0]['folder']=='invalid':
                self.pool.get('poweremail.core_accounts').write(cr,uid,context['active_ids'][0],{'isfolder':self.read(cr,uid,ids,['folder'])[0]['folder']})
                return {'type':'ir.actions.act_window_close'}
            else:
                raise osv.except_osv(_("Folder Error"), _("This is an invalid folder "))
        else:
            raise osv.except_osv(_("Folder Error"), _("Select a folder before you save record "))
        
poweremail_core_selfolder()
