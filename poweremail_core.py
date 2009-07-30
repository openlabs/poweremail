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
        'name': fields.char('Email Account Desc', size=64, required=True, readonly=True, select=True, states={'draft':[('readonly',False)]} ),
        'user':fields.many2one('res.users','Related User',required=True,readonly=True, states={'draft':[('readonly',False)]} ),
        
        'email_id': fields.char('Email ID',size=120,required=True, readonly=True, states={'draft':[('readonly',False)]} , help=" eg:yourname@yourdomain.com "),
        
        'smtpserver': fields.char('Server', size=120, required=True, readonly=True, states={'draft':[('readonly',False)]}, help="Enter name of outgoing server,eg:smtp.gmail.com " ),
        'smtpport': fields.integer('SMTP Port ', size=64, required=True, readonly=True, states={'draft':[('readonly',False)]}, help="Enter port number,eg:SMTP-587 "),
        'smtpuname': fields.char('User Name', size=120, required=True, readonly=True, states={'draft':[('readonly',False)]}),
        'smtppass': fields.char('Password', size=120, invisible=True, required=True, readonly=True, states={'draft':[('readonly',False)]}),
        'smtpssl':fields.boolean('Use SSL', states={'draft':[('readonly',False)]}),
        
        'iserver':fields.char('Incoming Server',size=100, readonly=True, states={'draft':[('readonly',False)]}, help="Enter name of incoming server,eg:imap.gmail.com "),
        'isport': fields.integer('Port', readonly=True, states={'draft':[('readonly',False)]}, help="For example IMAP: 993,POP3:995 "),
        'isuser':fields.char('User Name',size=100, readonly=True, states={'draft':[('readonly',False)]}),
        'ispass':fields.char('Password',size=100, readonly=True, states={'draft':[('readonly',False)]}),
        'iserver_type': fields.selection([('imap','IMAP'),('pop3','POP3')], 'Server Type',readonly=True, states={'draft':[('readonly',False)]}),
        'isssl':fields.boolean('Use SSL', readonly=True, states={'draft':[('readonly',False)]} ),
        'isfolder':fields.char('Folder',readonly=True,size=100,help="Folder to be used for downloading IMAP mails\nClick on adjacent button to select from a list of folders"),
        
        'company':fields.selection([
                                    ('yes','Yes'),
                                    ('no','No')
                                    ],'Company Mail A/c', readonly=True, help="Select if this mail account does not belong to specific user but the organisation as a whole. eg:info@somedomain.com",required=True, states={'draft':[('readonly',False)]}),

        'state':fields.selection([
                                  ('draft','Initiated'),
                                  ('suspended','Suspended'),
                                  ('approved','Approved')
                                  ],'Account Status',required=True,readonly=True),
                }

    _defaults = {
         'name':lambda self,cr,uid,ctx:self.pool.get('res.users').read(cr,uid,uid,['name'])['name'],
         'smtpssl':lambda *a:True,
         'state':lambda *a:'draft',
         'user':lambda self,cr,uid,ctx:uid,
         'iserver_type':lambda *a:'imap',
         'isssl': lambda *a: True,
         
                 }
                 
    _sql_constraints = [
        ('email_uniq', 'unique (email_id)', 'Another setting already exists with this email ID !')
    ]
    def _constraint_unique(self, cr, uid, ids):
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
    
    def on_change_emailid(self, cr, uid, ids, name=None,email_id=None, context=None):
        self.write(cr, uid, ids, {'state':'draft'}, context=context)
        return {'value': {'state': 'draft','smtpuname':email_id,'isuser':email_id}}
    
    
    def out_connection(self,cr,uid,ids,context={}):
        rec = self.browse(cr, uid, ids )[0]
        if rec:
            if rec.smtpserver and rec.smtpport and rec.smtpuname and rec.smtppass:
                try:
                    serv = smtplib.SMTP(rec.smtpserver,rec.smtpport)
                    if rec.smtpssl:
                        serv.ehlo()
                        serv.starttls()
                        serv.ehlo()
                except Exception,error:
                    raise osv.except_osv(_("SMTP Server Error"), _("An error occurred : %s ") % error)
                try:
                    serv.login(rec.smtpuname, rec.smtppass)
                except Exception,error:
                    raise osv.except_osv(_("SMTP Server Login Error"), _("An error occurred : %s ") % error)
                raise osv.except_osv(_("Information"),_("SMTP Test Connection Was Successful"))

    def in_connection(self,cr,uid,ids,context={}):
        rec = self.browse(cr, uid, ids )[0]
        if rec:
            if rec.iserver and rec.isport and rec.isuser and rec.ispass:
                if rec.iserver_type =='imap':
                    try:
                        if rec.isssl:
                            serv = imaplib.IMAP4_SSL(rec.iserver,rec.isport)
                        else:
                            serv = imaplib.IMAP4(rec.iserver,rec.isport)
                    except imaplib.IMAP4.error,error:
                        raise osv.except_osv(_("IMAP Server Error"), _("An error occurred : %s ") % error)
                    try:
                        serv.login(rec.isuser, rec.ispass)
                    except imaplib.IMAP4.error,error:
                        raise osv.except_osv(_("IMAP Server Login Error"), _("An error occurred : %s ") % error)
                    raise osv.except_osv(_("Information"),_("IMAP Test Connection Was Successful"))
                else:
                    try:
                        if rec.isssl:
                            serv = poplib.POP3_SSL(rec.iserver,rec.isport)
                        else:
                            serv = poplib.POP3(rec.iserver,rec.isport)
                    except Exception,error:
                        raise osv.except_osv(_("POP3 Server Error"), _("An error occurred : %s ") % error)
                    try:
                        serv.user(rec.isuser)
                        serv.pass_(rec.ispass)
                    except Exception,error:
                        raise osv.except_osv(_("POP3 Server Login Error"), _("An error occurred : %s ") % error)
                    raise osv.except_osv(_("Information"),_("POP3 Test Connection Was Successful"))

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

#**************************** MAIL SENDING FEATURES ***********************#
    def send_mail(self,cr,uid,ids,to=[],cc=[],bcc=[],subject="",body_text="",body_html="",payload={}):
        #ids:(from) Account from which mail is to be sent
        #to: To ids as list
        #cc: CC ids as list
        #bcc: BCC ids as list
        #subject: Subject as string
        #body_text: body as plain text
        #body_html: body in html
        #payload: attachments as binary dic. Eg:payload={'filename1.pdf':<binary>,'filename2.jpg':<binary>}
        #################### For each account in chosen accounts ##################
        logger = netsvc.Logger()
        for id in ids:
            core_obj = self.browse(cr,uid,id)
            if core_obj.smtpserver and core_obj.smtpport and core_obj.smtpuname and core_obj.smtppass and core_obj.state=='approved':
                try:
                    serv = smtplib.SMTP(core_obj.smtpserver,core_obj.smtpport)
                    if core_obj.smtpssl:
                        serv.ehlo()
                        serv.starttls()
                        serv.ehlo()
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Could not connect to server\nError: %s")% (id,error))
                    return False
                try:
                    serv.login(core_obj.smtpuname, core_obj.smtppass)
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Could not login to server\nError: %s")% (id,error))
                    return False
                try:
                    msg = MIMEMultipart('alternative')
                    msg['Subject']= subject
                    msg['From'] = str(core_obj.name + "<" + core_obj.email_id + ">")
                    msg['To']=",".join(map(str, to))
                    msg['cc']=",".join(map(str, cc)) or False
                    msg['bcc']=",".join(map(str, bcc)) or False
                    # Record the MIME types of both parts - text/plain and text/html.
                    part1 = MIMEText(body_text, 'plain')
                    if body_html:#If html body also exists, send that
                        part2 = MIMEText(body_html, 'html')
                    else:
                        part2 = part1
                    # Attach parts into message container.
                    # According to RFC 2046, the last part of a multipart message, in this case
                    # the HTML message, is best and preferred.
                    msg.attach(part1)
                    msg.attach(part2)
                    #Now add attachments if any
                    for file in payload.keys():
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(payload[file])
                        Encoders.encode_base64(part)
                        part.add_header('Content-Disposition', 'attachment; filename="%s"' % file)
                        msg.attach(part)
                    #msg is now complete, send it to everybody
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:MIME Error\nDescription: %s")% (id,error))
                    return False
                try:
                    serv.sendmail(msg['From'],to+cc+bcc)
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Server Send Error\nDescription: %s")% (id,error))
                    return False
                #The mail sending is complete
                serv.quit()
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s successfully Sent.")% (id))
                return True
            else:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Account not approved")% id)
                return False

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
