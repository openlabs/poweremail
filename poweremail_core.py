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
from html2text import html2text
import re
import smtplib
import base64
from email import Encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import netsvc
import poplib
import imaplib
import string
import email
import time
import poweremail_engines

class poweremail_core_accounts(osv.osv):
    _name = "poweremail.core_accounts"
    _known_content_types = ['multipart/mixed',
                            'multipart/alternative',
                            'multipart/related',
                            'text/plain',
                            'text/html'
                            ]
    _columns = {
        'name': fields.char('Email Account Desc', size=64, required=True, readonly=True, select=True, states={'draft':[('readonly',False)]} ),
        'user':fields.many2one('res.users','Related User',required=True,readonly=True, states={'draft':[('readonly',False)]} ),
        
        'email_id': fields.char('Email ID',size=120,required=True, readonly=True, states={'draft':[('readonly',False)]} , help=" eg:yourname@yourdomain.com "),
        
        'smtpserver': fields.char('Server', size=120, required=True, readonly=True, states={'draft':[('readonly',False)]}, help="Enter name of outgoing server,eg:smtp.gmail.com " ),
        'smtpport': fields.integer('SMTP Port ', size=64, required=True, readonly=True, states={'draft':[('readonly',False)]}, help="Enter port number,eg:SMTP-587 "),
        'smtpuname': fields.char('User Name', size=120, required=True, readonly=True, states={'draft':[('readonly',False)]}),
        'smtppass': fields.char('Password', size=120, invisible=True, required=True, readonly=True, states={'draft':[('readonly',False)]}),
        'smtpssl':fields.boolean('Use SSL', states={'draft':[('readonly',False)]}, readonly=True),
        'send_pref':fields.selection([
                                      ('html','HTML otherwise Text'),
                                      ('text','Text otherwise HTML'),
                                      ('both','Both HTML & Text')
                                      ],'Mail Format',required=True),
        'iserver':fields.char('Incoming Server',size=100, readonly=True, states={'draft':[('readonly',False)]}, help="Enter name of incoming server,eg:imap.gmail.com "),
        'isport': fields.integer('Port', readonly=True, states={'draft':[('readonly',False)]}, help="For example IMAP: 993,POP3:995 "),
        'isuser':fields.char('User Name',size=100, readonly=True, states={'draft':[('readonly',False)]}),
        'ispass':fields.char('Password',size=100, readonly=True, states={'draft':[('readonly',False)]}),
        'iserver_type': fields.selection([('imap','IMAP'),('pop3','POP3')], 'Server Type',readonly=True, states={'draft':[('readonly',False)]}),
        'isssl':fields.boolean('Use SSL', readonly=True, states={'draft':[('readonly',False)]} ),
        'isfolder':fields.char('Folder',readonly=True,size=100,help="Folder to be used for downloading IMAP mails\nClick on adjacent button to select from a list of folders"),
        'last_mail_id':fields.integer('Last Downloaded Mail',readonly=True),
        'rec_headers_den_mail':fields.boolean('First Receive headers, then download mail'),
        'dont_auto_down_attach':fields.boolean('Dont Download attachments automatically'),
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
         'last_mail_id':lambda *a:0,
         'rec_headers_den_mail':lambda *a:True,
         'dont_auto_down_attach':lambda *a:True,
         'send_pref':lambda *a: 'html',
                 }
                 
    _sql_constraints = [
        ('email_uniq', 'unique (email_id)', 'Another setting already exists with this email ID !')
    ]
    def _constraint_unique(self, cr, uid, ids):
        if self.read(cr,uid,ids,['company'])[0]['company']=='no':
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
        #checks SMTP credentials and confirms if outgoing connection works
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
        #Checks IMAP or POP3 credentials and returns if credentials are right
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

    def smtp_connection(self,cr,uid,id):
        #This function returns a SMTP server object
        logger = netsvc.Logger()
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
            #Everything is complete, now return the connection
            return serv
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Account not approved")% id)
            return False
                      
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
            serv = self.smtp_connection(cr, uid, id)
            if serv:
                try:
                    msg = MIMEMultipart()
                    if subject:
                        msg['Subject']= subject
                    msg['From'] = unicode(core_obj.name + "<" + core_obj.email_id + ">")
                    toadds = []
                    if to:
                        while (type(to)==type([])) and (False in to):
                            to = to.remove(False)
                        if (type(to)==type([])):
                            msg['To'] = ",".join(map(unicode,to))
                            toadds = to
                    if cc:
                        while (type(cc)==type([])) and (False in cc):
                            cc = cc.remove(False)
                        if (type(cc)==type([])):
                            msg['CC'] = ",".join(map(unicode,cc))
                            toadds += cc
                    if bcc:
                        while (type(bcc)==type([])) and (False in bcc):
                            bcc = bcc.remove(False)
                        #print "BCCCCCE:",bcc
                        if (type(bcc)==type([])):
                            #msg['BCC'] = ",".join(map(unicode,bcc)) #Dont show somebody gets a BCC
                            toadds += bcc
                    # Record the MIME types of both parts - text/plain and text/html.
                    if body_text:
                        l= body_text.replace(' ','')
                        l= l.replace('\r','')
                        l= l.replace('\n','')
                        l = len(l)
                        if l == 0:
                            body_text = False
                            
                    if not body_text:
                        if body_html:
                            body_text=html2text(body_html)
                        else:
                            body_text="Mail without body"
                    if not body_html:
                        body_html=body_text
                    # Attach parts into message container.
                    # According to RFC 2046, the last part of a multipart message, in this case
                    # the HTML message, is best and preferred.
                    if core_obj.send_pref == 'text' or core_obj.send_pref == 'both':
                        msg.attach(MIMEText(body_text, _charset='UTF-8'))
                    if core_obj.send_pref == 'html' or core_obj.send_pref == 'both':
                        msg.attach(MIMEText(body_html, _subtype='html',_charset='UTF-8'))
    
                    #Now add attachments if any
                    for file in payload.keys():
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(base64.decodestring(payload[file]))
                        f = open(file,"wb")
                        f.write(base64.decodestring(payload[file]))
                        f.close()
                        part.add_header('Content-Disposition', 'attachment; filename="%s"' % file)
                        Encoders.encode_base64(part)
                        msg.attach(part)
                        #msg is now complete, send it to everybody
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:MIME Error\nDescription: %s")% (id,error))
                    return False
                try:
                    #print msg['From'],toadds
                    serv.sendmail(unicode(msg['From']),toadds,msg.as_string())
                except Exception,error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Server Send Error\nDescription: %s")% (id,error))
                    return False
                #The mail sending is complete
                serv.close()
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Mail from Account %s successfully Sent.")% (id))
                return True
            else:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Account not approved")% id)
                return False
    
    def save_header(self,cr,uid,mail,coreaccountid,serv_ref):
        #Internal function for saving of mail headers to mailbox
        #mail: eMail Object
        #coreaccounti: ID of poeremail core account
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')
        
        vals = {
            'pem_from':mail['From'],
            'pem_to':mail['To'] or 'no recepient',
            'pem_cc':mail['cc'],
            'pem_bcc':mail['bcc'],
            'pem_recd':mail['date'],
            'date_mail':time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':mail['subject'],
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':'unread',
            'pem_body_text':'Mail not downloaded...',
            'pem_body_html':'Mail not downloaded...',
            'pem_account_id':coreaccountid
            }
        #Identify Mail Type
        if mail.get_content_type() in self._known_content_types:
            vals['mail_type']=mail.get_content_type()
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_WARNING, _("Saving Header of unknown payload (%s) Account:%s.")% (mail.get_content_type(),coreaccountid))
        #Create mailbox entry in Mail
        try:
        #print vals
            crid = mail_obj.create(cr,uid,vals)
        except Exception,e:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Save Header->Mailbox create error Account:%s,Mail:%s")% (coreaccountid,serv_ref))
        #Check if a create was success
        if crid:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Header for Mail %s Saved successfully as ID:%s for Account:%s.")% (serv_ref,crid,coreaccountid))
            crid=False
            return True
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Mail->Mailbox create error Account:%s,Mail:%s")% (coreaccountid,serv_ref))

    def save_fullmail(self,cr,uid,mail,coreaccountid,serv_ref):
        #Internal function for saving of mails to mailbox
        #mail: eMail Object
        #coreaccounti: ID of poeremail core account
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')
        #TODO:If multipart save attachments and save ids 
        vals = {
            'pem_from':mail['From'],
            'pem_to':mail['To'],
            'pem_cc':mail['cc'],
            'pem_bcc':mail['bcc'],
            'pem_recd':mail['date'],
            'date_mail':time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':mail['subject'],
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':'unread',
            'pem_body_text':'Mail not downloaded...',#TODO:Replace with mail text
            'pem_body_html':'Mail not downloaded...',#TODO:Replace
            'pem_account_id':coreaccountid
            }
        parsed_mail = self.get_payloads(mail)
        vals['pem_body_text']=parsed_mail['text']
        vals['pem_body_html']=parsed_mail['html']
        #Create the mailbox item now
        try:
            crid = mail_obj.create(cr,uid,vals)
        except Exception,e:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Save Header->Mailbox create error Account:%s,Mail:%s")% (coreaccountid,serv_ref))
        #Check if a create was success
        if crid:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Header for Mail %s Saved successfully as ID:%s for Account:%s.")% (serv_ref,crid,coreaccountid))
            #If there are attachments save them as well
            if parsed_mail['attachments']:
                self.save_attachments(self,cr,uid,mail,crid,parsed_mail,coreaccountid)
            crid=False
            return True
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Mail->Mailbox create error Account:%s,Mail:%s")% (coreaccountid,mail[0].split()[0]))

    def complete_mail(self,cr,uid,mail,coreaccountid,serv_ref,mailboxref):
        #Internal function for saving of mails to mailbox
        #mail: eMail Object
        #coreaccountid: ID of poeremail core account
        #serv_ref:Mail ID in the IMAP/POP server
        #mailboxref: ID of record in malbox to complete
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')
        #TODO:If multipart save attachments and save ids
        vals = {
            'pem_from':mail['From'],
            'pem_to':mail['To'] or 'no recepient',
            'pem_cc':mail['cc'],
            'pem_bcc':mail['bcc'],
            'pem_recd':mail['date'],
            'date_mail':time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':mail['subject'],
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':'unread',
            'pem_body_text':'Mail not downloaded...',#TODO:Replace with mail text
            'pem_body_html':'Mail not downloaded...',#TODO:Replace
            'pem_account_id':coreaccountid
            }
        #Identify Mail Type and get payload
        parsed_mail = self.get_payloads(mail)
        vals['pem_body_text'] = unicode(parsed_mail['text'])
        vals['pem_body_html'] = unicode(parsed_mail['html'])
        #Create the mailbox item now
        crid=False
        try:
            crid = mail_obj.write(cr,uid,mailboxref,vals)
        except Exception,e:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Save Mail->Mailbox write error Account:%s,Mail:%s")% (coreaccountid,serv_ref))
        #Check if a create was success
        if crid:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Mail %s Saved successfully as ID:%s for Account:%s.")% (serv_ref,crid,coreaccountid))
            #If there are attachments save them as well
            if parsed_mail['attachments']:
                self.save_attachments(cr,uid,mail,mailboxref,parsed_mail,coreaccountid)
            return True
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Mail->Mailbox create error Account:%s,Mail:%s")% (coreaccountid,mail[0].split()[0]))

    def save_attachments(self,cr,uid,mail,id,parsed_mail,coreaccountid):
        logger = netsvc.Logger()
        att_obj = self.pool.get('ir.attachment')
        mail_obj = self.pool.get('poweremail.mailbox')
        att_ids = []
        for each in parsed_mail['attachments']:#Get each attachment
            new_att_vals={
                        'name':mail['subject'] + '(' + each[0] + ')',
                        'datas':base64.b64encode(each[2]),
                        'datas_fname':each[1],
                        'description':mail['subject'] + " [Type:" + each[0] + ", Filename:" + each[1] + "]",
                        'res_model':'poweremail.mailbox',
                        'res_id':id
                            }
            att_ids.append(att_obj.create(cr,uid,new_att_vals))
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Downloaded & saved %s attachments Account:%s.")% (len(att_ids),coreaccountid))
            #Now attach the attachment ids to mail
            if mail_obj.write(cr,uid,id,{'pem_attachments_ids':[[6, 0, att_ids]]}):
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Attachment to mail for %s relation success! Account:%s.")% (id,coreaccountid))
            else:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Attachment to mail for %s relation failed Account:%s.")% (id,coreaccountid))
                        
    def get_mails(self,cr,uid,ids,ctx={}):
        #The function downloads the mails from the POP3 or IMAP server
        #The headers/full mail download depends on settings in the account
        #IDS should be list of id of poweremail_coreaccounts
        logger = netsvc.Logger()
        #The Main reception function starts here
        for id in ids:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Starting Header reception for account:%s.")% (id))
            rec = self.browse(cr, uid, id )
            if rec:
                if rec.iserver and rec.isport and rec.isuser and rec.ispass :
                    if rec.iserver_type =='imap' and rec.isfolder:
                        #Try Connecting to Server
                        try:
                            if rec.isssl:
                                serv = imaplib.IMAP4_SSL(rec.iserver,rec.isport)
                            else:
                                serv = imaplib.IMAP4(rec.iserver,rec.isport)
                        except imaplib.IMAP4.error,error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Error Account:%s Error:%s.")% (id,error))
                        #Try logging in to server
                        try:
                            serv.login(rec.isuser, rec.ispass)
                        except imaplib.IMAP4.error,error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Login Error Account:%s Error:%s.")% (id,error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Server Connected & logged in successfully Account:%s.")% (id))
                        #Select IMAP folder
                        try:
                            typ,msg_count = serv.select(rec.isfolder)
                        except imaplib.IMAP4.error,error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Folder Selection Error Account:%s Error:%s.")% (id,error))
                            raise osv.osv_except(_('Power Email'),_('IMAP Server Folder Selection Error Account:%s Error:%s.\nCheck account settings if you have selected a folder.')% (id,error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder selected successfully Account:%s.")% (id))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder Statistics for Account:%s:%s")% (id,serv.status(rec.isfolder,'(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)')[1][0]))
                        #If there are newer mails than the ones in mailbox
                        #print int(msg_count[0]),rec.last_mail_id
                        if rec.last_mail_id < int(msg_count[0]):
                            if rec.rec_headers_den_mail:
                                #Download Headers Only
                                for i in range(rec.last_mail_id+1,int(msg_count[0])+1):
                                    typ,msg = serv.fetch(str(i),'(BODY.PEEK[HEADER])')
                                    for mails in msg:
                                        if type(mails)==type(('tuple','type')):
                                            mail = email.message_from_string(mails[1])
                                            if self.save_header(cr,uid,mail,id,mails[0].split()[0]):#If saved succedfully then increment last mail recd
                                                self.write(cr,uid,id,{'last_mail_id':mails[0].split()[0]})
                            else:#Receive Full Mail first time itself
                                #Download Full RF822 Mails
                                for i in range(rec.last_mail_id+1,int(msg_count[0])+1):
                                    typ,msg = serv.fetch(str(i),'(RFC822)')
                                    for mails in msg:
                                        if type(mails)==type(('tuple','type')):
                                            mail = email.message_from_string(mails[1])
                                            if self.save_fullmail(cr,uid,mail,id,mails[0].split()[0]):#If saved succedfully then increment last mail recd
                                                self.write(cr,uid,id,{'last_mail_id':mails[0].split()[0]})
                        serv.close()
                        serv.logout()
                    elif rec.iserver_type =='pop3':
                        #Try Connecting to Server
                        try:
                            if rec.isssl:
                                serv = poplib.POP3_SSL(rec.iserver,rec.isport)
                            else:
                                serv = poplib.POP3(rec.iserver,rec.isport)
                        except Exception,error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Error Account:%s Error:%s.")% (id,error))
                        #Try logging in to server
                        try:
                            serv.user(rec.isuser)
                            serv.pass_(rec.ispass)
                        except Exception,error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Login Error Account:%s Error:%s.")% (id,error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Server Connected & logged in successfully Account:%s.")% (id))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Statistics:%s mails of %s size for Account:%s")% (serv.stat()[0],serv.stat()[1],id))
                        #If there are newer mails than the ones in mailbox
                        if rec.last_mail_id < serv.stat()[0]:
                            if rec.rec_headers_den_mail:
                                #Download Headers Only
                                for msgid in range(rec.last_mail_id+1,serv.stat()[0]+1):
                                    resp,msg,octet = serv.top(msgid,20) #20 Lines from the content
                                    mail = email.message_from_string(string.join(msg,"\n"))
                                    if self.save_header(cr,uid,mail,id,msgid):#If saved succedfully then increment last mail recd
                                        self.write(cr,uid,id,{'last_mail_id':msgid})
                            else:#Receive Full Mail first time itself
                                #Download Full RF822 Mails
                                for msgid in range(rec.last_mail_id+1,serv.stat()[0]+1):
                                    resp,msg,octet = serv.retr(msgid) #Full Mail
                                    mail = email.message_from_string(string.join(msg,"\n"))
                                    if self.save_header(cr,uid,mail,id,msgid):#If saved succedfully then increment last mail recd
                                        self.write(cr,uid,id,{'last_mail_id':msgid})
                    else:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Incoming server login attempt dropped Account:%s Check if Incoming server attributes are complete.")% (id))

    def get_fullmail(self,cr,uid,mailid,ctx):
        #The function downloads the full mail for which only header was downloaded
        #ID:of poeremail core account
        #ctx : should have mailboxref, the ID of mailbox record
        server_ref = self.pool.get('poweremail.mailbox').read(cr,uid,mailid,['server_ref'])['server_ref']
        id = self.pool.get('poweremail.mailbox').read(cr,uid,mailid,['pem_account_id'])['pem_account_id'][0]
        logger = netsvc.Logger()
        #The Main reception function starts here
        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Starting Full mail reception for mail:%s.")% (id))
        rec = self.browse(cr, uid, id )
        if rec:
            if rec.iserver and rec.isport and rec.isuser and rec.ispass :
                if rec.iserver_type =='imap' and rec.isfolder:
                    #Try Connecting to Server
                    try:
                        if rec.isssl:
                            serv = imaplib.IMAP4_SSL(rec.iserver,rec.isport)
                        else:
                            serv = imaplib.IMAP4(rec.iserver,rec.isport)
                    except imaplib.IMAP4.error,error:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Error Account:%s Error:%s.")% (id,error))
                    #Try logging in to server
                    try:
                        serv.login(rec.isuser, rec.ispass)
                    except imaplib.IMAP4.error,error:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Login Error Account:%s Error:%s.")% (id,error))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Server Connected & logged in successfully Account:%s.")% (id))
                    #Select IMAP folder
                    try:
                        typ,msg_count = serv.select(rec.isfolder)#typ,msg_count: practically not used here
                    except imaplib.IMAP4.error,error:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Folder Selection Error Account:%s Error:%s.")% (id,error))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder selected successfully Account:%s.")% (id))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder Statistics for Account:%s:%s")% (id,serv.status(rec.isfolder,'(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)')[1][0]))
                    #If there are newer mails than the ones in mailbox
                    typ,msg = serv.fetch(str(server_ref),'(RFC822)')
                    for mails in msg:
                        if type(mails)==type(('tuple','type')):
                            mail = email.message_from_string(mails[1])
                            self.complete_mail(cr,uid,mail,id,server_ref,mailid)
                    serv.close()
                    serv.logout()
                elif rec.iserver_type =='pop3':
                    #Try Connecting to Server
                    try:
                        if rec.isssl:
                            serv = poplib.POP3_SSL(rec.iserver,rec.isport)
                        else:
                            serv = poplib.POP3(rec.iserver,rec.isport)
                    except Exception,error:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Error Account:%s Error:%s.")% (id,error))
                    #Try logging in to server
                    try:
                        serv.user(rec.isuser)
                        serv.pass_(rec.ispass)
                    except Exception,error:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Login Error Account:%s Error:%s.")% (id,error))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Server Connected & logged in successfully Account:%s.")% (id))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Statistics:%s mails of %s size for Account:%s:%s")% (serv.stat()[0],serv.stat()[1],id))
                    #Download Full RF822 Mails
                    resp,msg,octet = serv.retr(server_ref) #Full Mail
                    mail = email.message_from_string(string.join(msg,"\n"))
                    self.complete_mail(cr,uid,mail,id,server_ref)
                else:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Incoming server login attempt dropped Account:%s Check if Incoming server attributes are complete.")% (id))
    
    def send_receive(self,cr,uid,ids,context={}):
        self.get_mails(cr, uid, ids, context)
        for id in ids:
            self.pool.get('poweremail.mailbox').send_all_mail(cr,uid,[],ctx={'filters':[('pem_account_id','=',id)]})
            
    def get_payloads(self,mail):
        #This function will go through the mail and identify the payloads and return them
        parsed_mail = {
                'text':False,
                'html':False,
                'attachments':[]
                       }
        for part in mail.walk():
            mail_part_type = part.get_content_type()
            if mail_part_type == 'text/plain':
                parsed_mail['text']=unicode(part.get_payload())
            elif mail_part_type == 'text/html':
                parsed_mail['html']=unicode(part.get_payload())
            elif not mail_part_type.startswith('multipart'):
                parsed_mail['attachments'].append((mail_part_type,part.get_filename(),part.get_payload(decode=True)))
        return parsed_mail

poweremail_core_accounts()


class poweremail_core_selfolder(osv.osv_memory):
    _name = "poweremail.core_selfolder"
    _description = "Shows a list of IMAP folders"

    def makereadable(self, imap_folder):
        if imap_folder:
            result = re.search(r'(?:\([^\)]*\)\s\")(.)(?:\"\s)(?:\")([^\"]*)(?:\")', imap_folder)
            seperator = result.groups()[0]
            folder_readable_name = ""
            splitname = result.groups()[1].split(seperator) #Not readable now
            if len(splitname) > 1:#If a parent and child exists, format it as parent/child/grandchild
                for i in range(0, len(splitname)-1):
                    folder_readable_name = splitname[i] + '/'
                folder_readable_name = folder_readable_name + splitname[-1]
            else:
                folder_readable_name = result.groups()[1].split(seperator)[0]
            return folder_readable_name
        return False
    
    def _get_folders(self, cr, uid, ctx={}):
        #print cr, uid, ctx
        if 'active_ids' in ctx.keys():
            record = self.pool.get('poweremail.core_accounts').browse(cr, uid, ctx['active_ids'][0])
            #print record.email_id
            if record:
                folderlist = []
                try:
                    if record.isssl:
                        serv = imaplib.IMAP4_SSL(record.iserver, record.isport)
                    else:
                        serv = imaplib.IMAP4(record.iserver, record.isport)
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Error"), _("An error occurred : %s ") % error)
                except Exception,error:
                    raise osv.except_osv(_("IMAP Server connection Error"), _("An error occurred : %s ") % error)
                try:
                    serv.login(record.isuser, record.ispass)
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Login Error"), _("An error occurred : %s ") % error)
                except Exception,error:
                    raise osv.except_osv(_("IMAP Server Login Error"), _("An error occurred : %s ") % error)
                try:
                    for folders in serv.list()[1]:
                        folder_readable_name = self.makereadable(folders)
                        if folders.find('Noselect') == -1: #If it is a selectable folder
                            folderlist.append((folder_readable_name, folder_readable_name))
                        if folder_readable_name == 'INBOX':
                            self.inboxvalue = folder_readable_name
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Folder Error"), _("An error occurred : %s ") % error)
                except Exception,error:
                    raise osv.except_osv(_("IMAP Server Folder Error"), _("An error occurred : %s ") % error)
            else:
                folderlist = [('invalid', 'Invalid')]
        else:
            folderlist = [('invalid', 'Invalid')]
        return folderlist

    _columns = {
        'name':fields.many2one('poweremail.core_accounts', string='Email Account', readonly=True),
        'folder':fields.selection(_get_folders, string="IMAP Folder"),        
    }

    _defaults = {
        'name':lambda self, cr, uid, ctx: ctx['active_ids'][0],
        'folder': lambda self, cr, uid, ctx:self.inboxvalue
    }

    def sel_folder(self, cr, uid, ids, context={}):
        if self.read(cr, uid, ids, ['folder'])[0]['folder']:
            if not self.read(cr, uid, ids, ['folder'])[0]['folder'] == 'invalid':
                self.pool.get('poweremail.core_accounts').write(cr, uid, context['active_ids'][0], {'isfolder':self.read(cr, uid, ids, ['folder'])[0]['folder']})
                return {'type':'ir.actions.act_window_close'}
            else:
                raise osv.except_osv(_("Folder Error"), _("This is an invalid folder "))
        else:
            raise osv.except_osv(_("Folder Error"), _("Select a folder before you save record "))
        
poweremail_core_selfolder()