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
from email.header import decode_header, Header
from email.utils import formatdate
import re
import netsvc
import poplib
import imaplib
import string
import email
import time, datetime
import poweremail_engines
from tools.translate import _
import tools

class poweremail_core_accounts(osv.osv):
    """
    Object to store email account settings
    """
    _name = "poweremail.core_accounts"
    _known_content_types = ['multipart/mixed',
                            'multipart/alternative',
                            'multipart/related',
                            'text/plain',
                            'text/html'
                            ]
    _columns = {
        'name': fields.char('Email Account Desc',
                        size=64, required=True,
                        readonly=True, select=True,
                        states={'draft':[('readonly', False)]}),
        'user':fields.many2one('res.users',
                        'Related User', required=True,
                        readonly=True, states={'draft':[('readonly', False)]}),
        'email_id': fields.char('Email ID',
                        size=120, required=True,
                        readonly=True, states={'draft':[('readonly', False)]} ,
                        help="eg: yourname@yourdomain.com"),
        'smtpserver': fields.char('Server',
                        size=120, required=True,
                        readonly=True, states={'draft':[('readonly', False)]},
                        help="Enter name of outgoing server, eg: smtp.gmail.com"),
        'smtpport': fields.integer('SMTP Port ',
                        size=64, required=True,
                        readonly=True, states={'draft':[('readonly', False)]},
                        help="Enter port number, eg: SMTP-587 "),
        'smtpuname': fields.char('User Name',
                        size=120, required=False,
                        readonly=True, states={'draft':[('readonly', False)]}),
        'smtppass': fields.char('Password',
                        size=120, invisible=True,
                        required=False, readonly=True,
                        states={'draft':[('readonly', False)]}),
        'smtptls':fields.boolean('Use TLS',
                        states={'draft':[('readonly', False)]}, readonly=True),
        'smtpssl':fields.boolean('Use SSL/TLS (only in python 2.6)',
                        states={'draft':[('readonly', False)]}, readonly=True),
        'send_pref':fields.selection([
                                      ('html', 'HTML otherwise Text'),
                                      ('text', 'Text otherwise HTML'),
                                      ('both', 'Both HTML & Text')
                                      ], 'Mail Format', required=True),
        'iserver':fields.char('Incoming Server',
                        size=100, readonly=True,
                        states={'draft':[('readonly', False)]},
                        help="Enter name of incoming server, eg: imap.gmail.com"),
        'isport': fields.integer('Port',
                        readonly=True, states={'draft':[('readonly', False)]},
                        help="For example IMAP: 993,POP3:995"),
        'isuser':fields.char('User Name',
                        size=100, readonly=True,
                        states={'draft':[('readonly', False)]}),
        'ispass':fields.char('Password',
                        size=100, readonly=True,
                        states={'draft':[('readonly', False)]}),
        'iserver_type': fields.selection([
                        ('imap', 'IMAP'),
                        ('pop3', 'POP3')
                        ], 'Server Type', readonly=True,
                        states={'draft':[('readonly', False)]}),
        'isssl':fields.boolean('Use SSL',
                        readonly=True, states={
                                           'draft':[('readonly', False)]
                                           }),
        'isfolder':fields.char('Folder',
                        readonly=True, size=100,
                        help="Folder to be used for downloading IMAP mails.\n" \
                        "Click on adjacent button to select from a list of folders."),
        'last_mail_id':fields.integer(
                        'Last Downloaded Mail', readonly=True),
        'rec_headers_den_mail':fields.boolean(
                        'First Receive headers, then download mail'),
        'dont_auto_down_attach':fields.boolean(
                        'Dont Download attachments automatically'),
        'allowed_groups':fields.many2many(
                        'res.groups',
                        'account_group_rel', 'templ_id', 'group_id',
                        string="Allowed User Groups",
                        help="Only users from these groups will be " \
                        "allowed to send mails from this ID."),
        'company':fields.selection([
                        ('yes', 'Yes'),
                        ('no', 'No')
                        ], 'Company Mail A/c',
                        readonly=True,
                        help="Select if this mail account does not belong " \
                        "to specific user but the organisation as a whole. " \
                        "eg: info@somedomain.com",
                        required=True, states={
                                           'draft':[('readonly', False)]
                                           }),

        'state':fields.selection([
                                  ('draft', 'Initiated'),
                                  ('suspended', 'Suspended'),
                                  ('approved', 'Approved')
                                  ],
                        'Account Status', required=True, readonly=True),
    }

    _defaults = {
         'name':lambda self, cursor, user, context:self.pool.get(
                                                'res.users'
                                                ).read(
                                                        cursor,
                                                        user,
                                                        user,
                                                        ['name'],
                                                        context
                                                        )['name'],
         'smtpssl':lambda * a:True,
         'state':lambda * a:'draft',
         'user':lambda self, cursor, user, context:user,
         'iserver_type':lambda * a:'imap',
         'isssl': lambda * a: True,
         'last_mail_id':lambda * a:0,
         'rec_headers_den_mail':lambda * a:True,
         'dont_auto_down_attach':lambda * a:True,
         'send_pref':lambda * a: 'html',
         'smtptls':lambda * a:True,
     }

    _sql_constraints = [
        (
         'email_uniq',
         'unique (email_id)',
         'Another setting already exists with this email ID !')
    ]

    def _constraint_unique(self, cursor, user, ids):
        """
        This makes sure that you dont give personal
        users two accounts with same ID (Validated in sql constaints)
        However this constraint exempts company accounts.
        Any no of co accounts for a user is allowed
        """
        if self.read(cursor, user, ids, ['company'])[0]['company'] == 'no':
            accounts = self.search(cursor, user, [
                                                 ('user', '=', user),
                                                 ('company', '=', 'no')
                                                 ])
            if len(accounts) > 1 :
                return False
            else :
                return True
        else:
            return True

    _constraints = [
        (_constraint_unique,
         'Error: You are not allowed to have more than 1 account.',
         [])
    ]

    def on_change_emailid(self, cursor, user, ids, name=None, email_id=None, context=None):
        """
        Called when the email ID field changes.
        
        UI enhancement
        Writes the same email value to the smtpusername
        and incoming username
        """
        #TODO: Check and remove the write. Is it needed?
        self.write(cursor, user, ids, {'state':'draft'}, context=context)
        return {
                'value': {
                          'state': 'draft',
                          'smtpuname':email_id,
                          'isuser':email_id
                          }
                }

    def _get_outgoing_server(self, cursor, user, ids, context=None):
        """
        Returns the Out Going Connection (SMTP) object
        
        @attention: DO NOT USE except_osv IN THIS METHOD
        @param cursor: Database Cursor
        @param user: ID of current user
        @param ids: ID/list of ids of current object for
                    which connection is required
                    First ID will be chosen from lists
        @param context: Context
        
        @return: SMTP server object or Exception
        """
        #Type cast ids to integer
        if type(ids) == list:
            ids = ids[0]
        this_object = self.browse(cursor, user, ids, context)
        if this_object:
            if this_object.smtpserver and this_object.smtpport:
                try:
                    if this_object.smtpssl:
                        serv = smtplib.SMTP_SSL(this_object.smtpserver, this_object.smtpport)
                    else:
                        serv = smtplib.SMTP(this_object.smtpserver, this_object.smtpport)
                    if this_object.smtptls:
                        serv.ehlo()
                        serv.starttls()
                        serv.ehlo()
                except Exception, error:
                    raise error
                try:
                    if serv.has_extn('AUTH') or this_object.smtpuname or this_object.smtppass:
                        serv.login(this_object.smtpuname, this_object.smtppass)
                except Exception, error:
                    raise error
                return serv
            raise Exception(_("SMTP SERVER or PORT not specified"))
        raise Exception(_("Core connection for the given ID does not exist"))

    def check_outgoing_connection(self, cursor, user, ids, context=None):
        """
        checks SMTP credentials and confirms if outgoing connection works
        (Attached to button)
        @param cursor: Database Cursor
        @param user: ID of current user
        @param ids: list of ids of current object for
                    which connection is required
        @param context: Context
        """
        try:
            self._get_outgoing_server(cursor, user, ids, context)
            raise osv.except_osv(_("SMTP Test Connection Was Successful"), '')
        except osv.except_osv, success_message:
            raise success_message
        except Exception, error:
            raise osv.except_osv(
                                 _("Out going connection test failed"),
                                 _("Reason: %s") % error
                                 )

    def _get_imap_server(self, record):
        """
        @param record: Browse record of current connection
        @return: IMAP or IMAP_SSL object
        """
        if record:
            if record.isssl:
                serv = imaplib.IMAP4_SSL(record.iserver, record.isport)
            else:
                serv = imaplib.IMAP4(record.iserver, record.isport)
            #Now try to login
            serv.login(record.isuser, record.ispass)
            return serv
        raise Exception(
                        _("Programming Error in _get_imap_server method. "
                        "The record received is invalid.")
                        )

    def _get_pop3_server(self, record):
        """
        @param record: Browse record of current connection
        @return: POP3 or POP3_SSL object
        """
        if record:
            if record.isssl:
                serv = poplib.POP3_SSL(record.iserver, record.isport)
            else:
                serv = poplib.POP3(record.iserver, record.isport)
            #Now try to login
            serv.user(record.isuser)
            serv.pass_(record.ispass)
            return serv
        raise Exception(
                        _("Programming Error in _get_pop3_server method. "
                        "The record received is invalid.")
                        )

    def _get_incoming_server(self, cursor, user, ids, context=None):
        """
        Returns the Incoming Server object
        Could be IMAP/IMAP_SSL/POP3/POP3_SSL
        
        @attention: DO NOT USE except_osv IN THIS METHOD
        
        @param cursor: Database Cursor
        @param user: ID of current user
        @param ids: ID/list of ids of current object for
                    which connection is required
                    First ID will be chosen from lists
        @param context: Context
        
        @return: IMAP/POP3 server object or Exception
        """
        #Type cast ids to integer
        if type(ids) == list:
            ids = ids[0]
        this_object = self.browse(cursor, user, ids, context)
        if this_object:
            #First validate data
            if not this_object.iserver:
                raise Exception(_("Incoming server is not defined"))
            if not this_object.isport:
                raise Exception(_("Incoming port is not defined"))
            if not this_object.isuser:
                raise Exception(_("Incoming server user name is not defined"))
            if not this_object.isuser:
                raise Exception(_("Incoming server password is not defined"))
            #Now fetch the connection
            if this_object.iserver_type == 'imap':
                serv = self._get_imap_server(this_object)
            elif this_object.iserver_type == 'pop3':
                serv = self._get_pop3_server(this_object)
            return serv
        raise Exception(
                    _("The specified record for connection does not exist")
                        )

    def check_incoming_connection(self, cursor, user, ids, context=None):
        """
        checks incoming credentials and confirms if outgoing connection works
        (Attached to button)
        @param cursor: Database Cursor
        @param user: ID of current user
        @param ids: list of ids of current object for
                    which connection is required
        @param context: Context
        """
        try:
            self._get_incoming_server(cursor, user, ids, context)
            raise osv.except_osv(_("Incoming Test Connection Was Successful"), '')
        except osv.except_osv, success_message:
            raise success_message
        except Exception, error:
            raise osv.except_osv(
                                 _("In coming connection test failed"),
                                 _("Reason: %s") % error
                                 )

    def do_approval(self, cr, uid, ids, context={}):
        #TODO: Check if user has rights
        self.write(cr, uid, ids, {'state':'approved'}, context=context)
#        wf_service = netsvc.LocalService("workflow")

    def smtp_connection(self, cursor, user, id, context=None):
        """
        This method should now wrap smtp_connection
        """
        #This function returns a SMTP server object
        logger = netsvc.Logger()
        core_obj = self.browse(cursor, user, id, context)
        if core_obj.smtpserver and core_obj.smtpport and core_obj.state == 'approved':
            try:
                serv = self._get_outgoing_server(cursor, user, id, context)
            except Exception, error:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed on login. Probable Reason:Could not login to server\nError: %s") % (id, error))
                return False
            #Everything is complete, now return the connection
            return serv
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason: Account not approved") % id)
            return False

#**************************** MAIL SENDING FEATURES ***********************#
    def split_to_ids(self, ids_as_str):
        """
        Identifies email IDs separated by separators
        and returns a list
        TODO: Doc this
        "a@b.com,c@bcom; d@b.com;e@b.com->['a@b.com',...]"
        """
        email_sep_by_commas = ids_as_str \
                                    .replace('; ', ',') \
                                    .replace(';', ',') \
                                    .replace(', ', ',')
        return email_sep_by_commas.split(',')

    def get_ids_from_dict(self, addresses={}):
        """
        TODO: Doc this
        """
        result = {'all':[]}
        keys = ['To', 'CC', 'BCC']
        for each in keys:
            ids_as_list = self.split_to_ids(addresses.get(each, u''))
            while u'' in ids_as_list:
                ids_as_list.remove(u'')
            result[each] = ids_as_list
            result['all'].extend(ids_as_list)
        return result

    def send_mail(self, cr, uid, ids, addresses, subject='', body=None, payload=None, context=None):
        #TODO: Replace all this crap with a single email object
        if body is None:
            body = {}
        if payload is None:
            payload = {}
        if context is None:
            context = {}
        logger = netsvc.Logger()
        for id in ids:
            core_obj = self.browse(cr, uid, id, context)
            serv = self.smtp_connection(cr, uid, id)
            if serv:
                try:
                    msg = MIMEMultipart()
                    if subject:
                        msg['Subject'] = subject
                    sender_name = Header(core_obj.name, 'utf-8').encode()
                    msg['From'] = sender_name + " <" + core_obj.email_id + ">"
                    msg['Organization'] = tools.ustr(core_obj.user.company_id.name)
                    msg['Date'] = formatdate()
                    addresses_l = self.get_ids_from_dict(addresses)
                    if addresses_l['To']:
                        msg['To'] = u','.join(addresses_l['To'])
                    if addresses_l['CC']:
                        msg['CC'] = u','.join(addresses_l['CC'])
#                    if addresses_l['BCC']:
#                        msg['BCC'] = u','.join(addresses_l['BCC'])
                    if body.get('text', False):
                        temp_body_text = body.get('text', '')
                        l = len(temp_body_text.replace(' ', '').replace('\r', '').replace('\n', ''))
                        if l == 0:
                            body['text'] = u'No Mail Message'
                    # Attach parts into message container.
                    # According to RFC 2046, the last part of a multipart message, in this case
                    # the HTML message, is best and preferred.
                    if core_obj.send_pref == 'text' or core_obj.send_pref == 'both':
                        body_text = body.get('text', u'No Mail Message')
                        body_text = tools.ustr(body_text)
                        msg.attach(MIMEText(body_text.encode("utf-8"), _charset='UTF-8'))
                    if core_obj.send_pref == 'html' or core_obj.send_pref == 'both':
                        html_body = body.get('html', u'')
                        if len(html_body) == 0 or html_body == u'':
                            html_body = body.get('text', u'<p>No Mail Message</p>').replace('\n', '<br/>').replace('\r', '<br/>')
                        html_body = tools.ustr(html_body)
                        msg.attach(MIMEText(html_body.encode("utf-8"), _subtype='html', _charset='UTF-8'))
                    #Now add attachments if any
                    for file in payload.keys():
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(base64.decodestring(payload[file]))
                        part.add_header('Content-Disposition', 'attachment; filename="%s"' % file)
                        Encoders.encode_base64(part)
                        msg.attach(part)
                except Exception, error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason: MIME Error\nDescription: %s") % (id, error))
                    return error
                try:
                    #print msg['From'],toadds
                    serv.sendmail(msg['From'], addresses_l['all'], msg.as_string())
                except Exception, error:
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason:Server Send Error\nDescription: %s") % (id, error))
                    return error
                #The mail sending is complete
                serv.close()
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Mail from Account %s successfully Sent.") % (id))
                return True
            else:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Mail from Account %s failed. Probable Reason: Account not approved") % id)

    def extracttime(self, time_as_string):
        """
        TODO: DOC THis
        """
        logger = netsvc.Logger()
        #The standard email dates are of format similar to:
        #Thu, 8 Oct 2009 09:35:42 +0200
        #print time_as_string
        date_as_date = False
        convertor = {'+':1, '-':-1}
        try:
            time_as_string = time_as_string.replace(',', '')
            date_list = time_as_string.split(' ')
            date_temp_str = ' '.join(date_list[1:5])
            if len(date_list) >= 6:
                sign = convertor.get(date_list[5][0], False)
            else:
                sign = False
            try:
                dt = datetime.datetime.strptime(
                                            date_temp_str,
                                            "%d %b %Y %H:%M:%S")
            except:
                try:
                    dt = datetime.datetime.strptime(
                                            date_temp_str,
                                            "%d %b %Y %H:%M")
                except:
                    return False
            if sign:
                try:
                    offset = datetime.timedelta(
                                hours=sign * int(
                                             date_list[5][1:3]
                                                ),
                                             minutes=sign * int(
                                                            date_list[5][3:5]
                                                                )
                                                )
                except Exception, e2:
                    """Looks like UT or GMT, just forget decoding"""
                    return False
            else:
                offset = datetime.timedelta(hours=0)
            dt = dt + offset
            date_as_date = dt.strftime('%Y-%m-%d %H:%M:%S')
            #print date_as_date
        except Exception, e:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_WARNING,
                    _(
                      "Datetime Extraction failed. Date:%s " \
                      "\tError:%s") % (
                                    time_as_string,
                                    e)
                      )
        return date_as_date

    def save_header(self, cr, uid, mail, coreaccountid, serv_ref, context=None):
        """
        TODO:DOC this
        """
        if context is None:
            context = {}
        #Internal function for saving of mail headers to mailbox
        #mail: eMail Object
        #coreaccounti: ID of poeremail core account
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')

        vals = {
            'pem_from':self.decode_header_text(mail['From']),
            'pem_to':mail['To'] and self.decode_header_text(mail['To']) or 'no recepient',
            'pem_cc':self.decode_header_text(mail['cc']),
            'pem_bcc':self.decode_header_text(mail['bcc']),
            'pem_recd':mail['date'],
            'date_mail':self.extracttime(mail['date']) or time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':self.decode_header_text(mail['subject']),
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':context.get('state', 'unread'),
            'pem_body_text':'Mail not downloaded...',
            'pem_body_html':'Mail not downloaded...',
            'pem_account_id':coreaccountid
            }
        #Identify Mail Type
        if mail.get_content_type() in self._known_content_types:
            vals['mail_type'] = mail.get_content_type()
        else:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_WARNING,
                    _("Saving Header of unknown payload (%s) Account:%s.") % (
                                                      mail.get_content_type(),
                                                      coreaccountid)

                    )
        #Create mailbox entry in Mail
        crid = False
        try:
        #print vals
            crid = mail_obj.create(cr, uid, vals, context)
        except Exception, e:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_ERROR,
                    _("Save Header->Mailbox create error "
                    "Account:%s, Mail:%s, Error:%s") % (coreaccountid,
                                                     serv_ref, str(e)))
        #Check if a create was success
        if crid:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_INFO,
                    _("Header for Mail %s Saved successfully as "
                    "ID:%s for Account:%s.") % (serv_ref, crid, coreaccountid)
                    )
            crid = False
            return True
        else:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_ERROR,
                    _("IMAP Mail->Mailbox create error "
                    "Account:%s, Mail:%s") % (coreaccountid, serv_ref))

    def save_fullmail(self, cr, uid, mail, coreaccountid, serv_ref, context=None):
        """
        TODO: Doc this
        """
        if context is None:
            context = {}
        #Internal function for saving of mails to mailbox
        #mail: eMail Object
        #coreaccounti: ID of poeremail core account
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')
        #TODO:If multipart save attachments and save ids
        vals = {
            'pem_from':self.decode_header_text(mail['From']),
            'pem_to':self.decode_header_text(mail['To']),
            'pem_cc':self.decode_header_text(mail['cc']),
            'pem_bcc':self.decode_header_text(mail['bcc']),
            'pem_recd':mail['date'],
            'date_mail':self.extracttime(
                            mail['date']
                                ) or time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':self.decode_header_text(mail['subject']),
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':context.get('state', 'unread'),
            'pem_body_text':'Mail not downloaded...', #TODO:Replace with mail text
            'pem_body_html':'Mail not downloaded...', #TODO:Replace
            'pem_account_id':coreaccountid
            }
        parsed_mail = self.get_payloads(mail)
        vals['pem_body_text'] = parsed_mail['text']
        vals['pem_body_html'] = parsed_mail['html']
        #Create the mailbox item now
        crid = False
        try:
            crid = mail_obj.create(cr, uid, vals, context)
        except Exception, e:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_ERROR,
                    _("Save Header->Mailbox create error " \
                    "Account:%s, Mail:%s, Error:%s") % (
                                                    coreaccountid,
                                                    serv_ref,
                                                    str(e))
                                )
        #Check if a create was success
        if crid:
            logger.notifyChannel(
                    _("Power Email"),
                    netsvc.LOG_INFO,
                    _("Header for Mail %s Saved successfully " \
                    "as ID:%s for Account:%s.") % (serv_ref,
                                                  crid,
                                                  coreaccountid))
            #If there are attachments save them as well
            if parsed_mail['attachments']:
                self.save_attachments(cr, uid, mail, crid,
                                      parsed_mail, coreaccountid, context)
            crid = False
            return True
        else:
            logger.notifyChannel(
                                 _("Power Email"),
                                 netsvc.LOG_ERROR,
                                 _("IMAP Mail->Mailbox create error " \
                                 "Account:%s, Mail:%s") % (
                                                         coreaccountid,
                                                         mail[0].split()[0]))

    def complete_mail(self, cr, uid, mail, coreaccountid, serv_ref, mailboxref, context=None):
        if context is None:
            context = {}
        #Internal function for saving of mails to mailbox
        #mail: eMail Object
        #coreaccountid: ID of poeremail core account
        #serv_ref:Mail ID in the IMAP/POP server
        #mailboxref: ID of record in malbox to complete
        logger = netsvc.Logger()
        mail_obj = self.pool.get('poweremail.mailbox')
        #TODO:If multipart save attachments and save ids
        vals = {
            'pem_from':self.decode_header_text(mail['From']),
            'pem_to':mail['To'] and self.decode_header_text(mail['To']) or 'no recepient',
            'pem_cc':self.decode_header_text(mail['cc']),
            'pem_bcc':self.decode_header_text(mail['bcc']),
            'pem_recd':mail['date'],
            'date_mail':time.strftime("%Y-%m-%d %H:%M:%S"),
            'pem_subject':self.decode_header_text(mail['subject']),
            'server_ref':serv_ref,
            'folder':'inbox',
            'state':context.get('state', 'unread'),
            'pem_body_text':'Mail not downloaded...', #TODO:Replace with mail text
            'pem_body_html':'Mail not downloaded...', #TODO:Replace
            'pem_account_id':coreaccountid
            }
        #Identify Mail Type and get payload
        parsed_mail = self.get_payloads(mail)
        vals['pem_body_text'] = tools.ustr(parsed_mail['text'])
        vals['pem_body_html'] = tools.ustr(parsed_mail['html'])
        #Create the mailbox item now
        crid = False
        try:
            crid = mail_obj.write(cr, uid, mailboxref, vals, context)
        except Exception, e:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Save Mail->Mailbox write error Account:%s, Mail:%s") % (coreaccountid, serv_ref))
        #Check if a create was success
        if crid:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Mail %s Saved successfully as ID:%s for Account:%s.") % (serv_ref, crid, coreaccountid))
            #If there are attachments save them as well
            if parsed_mail['attachments']:
                self.save_attachments(cr, uid, mail, mailboxref, parsed_mail, coreaccountid, context)
            return True
        else:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Mail->Mailbox create error Account:%s, Mail:%s") % (coreaccountid, mail[0].split()[0]))

    def save_attachments(self, cr, uid, mail, id, parsed_mail, coreaccountid, context=None):
        logger = netsvc.Logger()
        att_obj = self.pool.get('ir.attachment')
        mail_obj = self.pool.get('poweremail.mailbox')
        att_ids = []
        for each in parsed_mail['attachments']:#Get each attachment
            new_att_vals = {
                        'name':self.decode_header_text(mail['subject']) + '(' + each[0] + ')',
                        'datas':base64.b64encode(each[2] or ''),
                        'datas_fname':each[1],
                        'description':(self.decode_header_text(mail['subject']) or _('No Subject')) + " [Type:" + (each[0] or 'Unknown') + ", Filename:" + (each[1] or 'No Name') + "]",
                        'res_model':'poweremail.mailbox',
                        'res_id':id
                            }
            att_ids.append(att_obj.create(cr, uid, new_att_vals, context))
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Downloaded & saved %s attachments Account:%s.") % (len(att_ids), coreaccountid))
            #Now attach the attachment ids to mail
            if mail_obj.write(cr, uid, id, {'pem_attachments_ids':[[6, 0, att_ids]]}, context):
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Attachment to mail for %s relation success! Account:%s.") % (id, coreaccountid))
            else:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Attachment to mail for %s relation failed Account:%s.") % (id, coreaccountid))

    def get_mails(self, cr, uid, ids, context=None):
        if context is None:
            context = {}
        #The function downloads the mails from the POP3 or IMAP server
        #The headers/full mail download depends on settings in the account
        #IDS should be list of id of poweremail_coreaccounts
        logger = netsvc.Logger()
        #The Main reception function starts here
        for id in ids:
            logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Starting Header reception for account:%s.") % (id))
            rec = self.browse(cr, uid, id, context)
            if rec:
                if rec.iserver and rec.isport and rec.isuser and rec.ispass :
                    if rec.iserver_type == 'imap' and rec.isfolder:
                        #Try Connecting to Server
                        try:
                            if rec.isssl:
                                serv = imaplib.IMAP4_SSL(rec.iserver, rec.isport)
                            else:
                                serv = imaplib.IMAP4(rec.iserver, rec.isport)
                        except imaplib.IMAP4.error, error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Error Account:%s Error:%s.") % (id, error))
                        #Try logging in to server
                        try:
                            serv.login(rec.isuser, rec.ispass)
                        except imaplib.IMAP4.error, error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Login Error Account:%s Error:%s.") % (id, error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Server Connected & logged in successfully Account:%s.") % (id))
                        #Select IMAP folder
                        try:
                            typ, msg_count = serv.select('"%s"' % rec.isfolder)
                        except imaplib.IMAP4.error, error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("IMAP Server Folder Selection Error Account:%s Error:%s.") % (id, error))
                            raise osv.except_osv(_('Power Email'), _('IMAP Server Folder Selection Error Account:%s Error:%s.\nCheck account settings if you have selected a folder.') % (id, error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder selected successfully Account:%s.") % (id))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("IMAP Folder Statistics for Account:%s:%s") % (id, serv.status('"%s"' % rec.isfolder, '(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)')[1][0]))
                        #If there are newer mails than the ones in mailbox
                        #print int(msg_count[0]),rec.last_mail_id
                        if rec.last_mail_id < int(msg_count[0]):
                            if rec.rec_headers_den_mail:
                                #Download Headers Only
                                for i in range(rec.last_mail_id + 1, int(msg_count[0]) + 1):
                                    typ, msg = serv.fetch(str(i), '(FLAGS BODY.PEEK[HEADER])')
                                    for mails in msg:
                                        if type(mails) == type(('tuple', 'type')):
                                            mail = email.message_from_string(mails[1])
                                            ctx = context.copy()
                                            if '\Seen' in mails[0]:
                                                ctx['state'] = 'read'
                                            if self.save_header(cr, uid, mail, id, mails[0].split()[0], ctx):#If saved succedfully then increment last mail recd
                                                self.write(cr, uid, id, {'last_mail_id':mails[0].split()[0]}, context)
                            else:#Receive Full Mail first time itself
                                #Download Full RF822 Mails
                                for i in range(rec.last_mail_id + 1, int(msg_count[0]) + 1):
                                    typ, msg = serv.fetch(str(i), '(FLAGS RFC822)')
                                    for j in range(0, len(msg) / 2):
                                        mails = msg[j * 2]
                                        flags = msg[(j * 2) + 1]
                                        if type(mails) == type(('tuple', 'type')):
                                            ctx = context.copy()
                                            if '\Seen' in flags:
                                                ctx['state'] = 'read'
                                            mail = email.message_from_string(mails[1])
                                            if self.save_fullmail(cr, uid, mail, id, mails[0].split()[0], ctx):#If saved succedfully then increment last mail recd
                                                self.write(cr, uid, id, {'last_mail_id':mails[0].split()[0]}, context)
                        serv.close()
                        serv.logout()
                    elif rec.iserver_type == 'pop3':
                        #Try Connecting to Server
                        try:
                            if rec.isssl:
                                serv = poplib.POP3_SSL(rec.iserver, rec.isport)
                            else:
                                serv = poplib.POP3(rec.iserver, rec.isport)
                        except Exception, error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Error Account:%s Error:%s.") % (id, error))
                        #Try logging in to server
                        try:
                            serv.user(rec.isuser)
                            serv.pass_(rec.ispass)
                        except Exception, error:
                            logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("POP3 Server Login Error Account:%s Error:%s.") % (id, error))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Server Connected & logged in successfully Account:%s.") % (id))
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Statistics:%s mails of %s size for Account:%s") % (serv.stat()[0], serv.stat()[1], id))
                        #If there are newer mails than the ones in mailbox
                        if rec.last_mail_id < serv.stat()[0]:
                            if rec.rec_headers_den_mail:
                                #Download Headers Only
                                for msgid in range(rec.last_mail_id + 1, serv.stat()[0] + 1):
                                    resp, msg, octet = serv.top(msgid, 20) #20 Lines from the content
                                    mail = email.message_from_string(string.join(msg, "\n"))
                                    if self.save_header(cr, uid, mail, id, msgid):#If saved succedfully then increment last mail recd
                                        self.write(cr, uid, id, {'last_mail_id':msgid}, context)
                            else:#Receive Full Mail first time itself
                                #Download Full RF822 Mails
                                for msgid in range(rec.last_mail_id + 1, serv.stat()[0] + 1):
                                    resp, msg, octet = serv.retr(msgid) #Full Mail
                                    mail = email.message_from_string(string.join(msg, "\n"))
                                    if self.save_header(cr, uid, mail, id, msgid):#If saved succedfully then increment last mail recd
                                        self.write(cr, uid, id, {'last_mail_id':msgid}, context)
                    else:
                        logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Incoming server login attempt dropped Account:%s Check if Incoming server attributes are complete.") % (id))

    def get_fullmail(self, cr, uid, mailid, context):
        #The function downloads the full mail for which only header was downloaded
        #ID:of poeremail core account
        #context: should have mailboxref, the ID of mailbox record
        server_ref = self.pool.get(
                        'poweremail.mailbox'
                        ).read(cr, uid, mailid,
                               ['server_ref'],
                               context)['server_ref']
        id = self.pool.get(
                        'poweremail.mailbox'
                        ).read(cr, uid, mailid,
                               ['pem_account_id'],
                               context)['pem_account_id'][0]
        logger = netsvc.Logger()
        #The Main reception function starts here
        logger.notifyChannel(
                _("Power Email"),
                netsvc.LOG_INFO,
                _("Starting Full mail reception for mail:%s.") % (id))
        rec = self.browse(cr, uid, id, context)
        if rec:
            if rec.iserver and rec.isport and rec.isuser and rec.ispass :
                if rec.iserver_type == 'imap' and rec.isfolder:
                    #Try Connecting to Server
                    try:
                        if rec.isssl:
                            serv = imaplib.IMAP4_SSL(
                                                     rec.iserver,
                                                     rec.isport
                                                     )
                        else:
                            serv = imaplib.IMAP4(
                                                 rec.iserver,
                                                 rec.isport
                                                 )
                    except imaplib.IMAP4.error, error:
                        logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_ERROR,
                                _(
                                  "IMAP Server Error Account:%s Error:%s."
                                  ) % (id, error))
                    #Try logging in to server
                    try:
                        serv.login(rec.isuser, rec.ispass)
                    except imaplib.IMAP4.error, error:
                        logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_ERROR,
                                _(
                        "IMAP Server Login Error Account:%s Error:%s."
                                ) % (id, error))
                    logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_INFO,
                                _(
                        "IMAP Server Connected & logged in " \
                        "successfully Account:%s."
                                ) % (id))
                    #Select IMAP folder
                    try:
                        typ, msg_count = serv.select('"%s"' % rec.isfolder)#typ,msg_count: practically not used here
                    except imaplib.IMAP4.error, error:
                        logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_ERROR,
                                _(
                      "IMAP Server Folder Selection " \
                      "Error Account:%s Error:%s."
                                  ) % (id, error))
                    logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_INFO,
                                _(
                      "IMAP Folder selected successfully Account:%s."
                                  ) % (id))
                    logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_INFO,
                                _(
                      "IMAP Folder Statistics for Account:%s:%s"
                                  ) % (
                           id,
                           serv.status(
                                '"%s"' % rec.isfolder,
                                '(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)'
                                )[1][0])
                                  )
                    #If there are newer mails than the ones in mailbox
                    typ, msg = serv.fetch(str(server_ref), '(FLAGS RFC822)')
                    for i in range(0, len(msg) / 2):
                        mails = msg[i * 2]
                        flags = msg[(i * 2) + 1]
                        if type(mails) == type(('tuple', 'type')):
                            if '\Seen' in flags:
                                context['state'] = 'read'
                            mail = email.message_from_string(mails[1])
                            self.complete_mail(cr, uid, mail, id,
                                               server_ref, mailid, context)
                    serv.close()
                    serv.logout()
                elif rec.iserver_type == 'pop3':
                    #Try Connecting to Server
                    try:
                        if rec.isssl:
                            serv = poplib.POP3_SSL(
                                            rec.iserver,
                                            rec.isport
                                                )
                        else:
                            serv = poplib.POP3(
                                            rec.iserver,
                                            rec.isport
                                            )
                    except Exception, error:
                        logger.notifyChannel(
                            _("Power Email"),
                            netsvc.LOG_ERROR,
                            _(
                        "POP3 Server Error Account:%s Error:%s."
                            ) % (id, error))
                    #Try logging in to server
                    try:
                        serv.user(rec.isuser)
                        serv.pass_(rec.ispass)
                    except Exception, error:
                        logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_ERROR,
                                _(
                    "POP3 Server Login Error Account:%s Error:%s."
                                ) % (id, error))
                    logger.notifyChannel(
                                _("Power Email"),
                                netsvc.LOG_INFO,
                                _(
                    "POP3 Server Connected & logged in " \
                    "successfully Account:%s."
                                ) % (id))
                    logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("POP3 Statistics:%s mails of %s size for Account:%s:") % (serv.stat()[0], serv.stat()[1], id))
                    #Download Full RF822 Mails
                    resp, msg, octet = serv.retr(server_ref) #Full Mail
                    mail = email.message_from_string(string.join(msg, "\n"))
                    self.complete_mail(cr, uid, mail, id,
                                       server_ref, mailid, context)
                else:
                    logger.notifyChannel(
                        _("Power Email"),
                        netsvc.LOG_ERROR,
                        _(
                        "Incoming server login attempt dropped Account:%s " \
                        "Check if Incoming server attributes are complete."
                        ) % (id))

    def send_receive(self, cr, uid, ids, context=None):
        self.get_mails(cr, uid, ids, context)
        for id in ids:
            ctx = context.copy()
            ctx['filters'] = [('pem_account_id', '=', id)]
            self.pool.get('poweremail.mailbox').send_all_mail(cr, uid, [], context=ctx)
        return True

    def get_payloads(self, mail):
        """
        """
        #This function will go through the mail and identify the payloads and return them
        parsed_mail = {
                'text':False,
                'html':False,
                'attachments':[]
                       }
        for part in mail.walk():
            mail_part_type = part.get_content_type()
            if mail_part_type == 'text/plain':
                parsed_mail['text'] = tools.ustr(part.get_payload(decode=True)) # decode=True to decode a MIME message
            elif mail_part_type == 'text/html':
                parsed_mail['html'] = tools.ustr(part.get_payload(decode=True)) # Is decode=True needed in html MIME messages?
            elif not mail_part_type.startswith('multipart'):
                parsed_mail['attachments'].append((mail_part_type, part.get_filename(), part.get_payload(decode=True)))
        return parsed_mail

    def decode_header_text(self, text):
        """ Decode internationalized headers RFC2822.
            To, CC, BCC, Subject fields can contain
            text slices with different encodes, like:
                =?iso-8859-1?Q?Enric_Mart=ED?= <enricmarti@company.com>,
                =?Windows-1252?Q?David_G=F3mez?= <david@company.com>
            Sometimes they include extra " character at the beginning/
            end of the contact name, like:
                "=?iso-8859-1?Q?Enric_Mart=ED?=" <enricmarti@company.com>
            and decode_header() does not work well, so we use regular
            expressions (?=   ? ?   ?=) to split the text slices
        """
        if not text:
            return text
        p = re.compile("(=\?.*?\?.\?.*?\?=)")
        text2 = ''
        try:
            for t2 in p.split(text):
                text2 += ''.join(
                            [s.decode(
                                      t or 'ascii'
                                    ) for (s, t) in decode_header(t2)]
                                ).encode('utf-8')
        except:
            return text
        return text2

poweremail_core_accounts()


class PoweremailSelectFolder(osv.osv_memory):
    _name = "poweremail.core_selfolder"
    _description = "Shows a list of IMAP folders"

    def makereadable(self, imap_folder):
        if imap_folder:
	    # We consider imap_folder may be in one of the following formats:
	    # A string like this: '(\HasChildren) "/" "INBOX"'
	    # Or a tuple like this: ('(\\HasNoChildren) "/" {18}', 'INBOX/contacts')
            if isinstance(imap_folder, tuple):
                return imap_folder[1]
            result = re.search(
                        r'(?:\([^\)]*\)\s\")(.)(?:\"\s)(?:\")?([^\"]*)(?:\")?',
                        imap_folder)
            seperator = result.groups()[0]
            folder_readable_name = ""
            splitname = result.groups()[1].split(seperator) #Not readable now
            #If a parent and child exists, format it as parent/child/grandchild
            if len(splitname) > 1:
                for i in range(0, len(splitname) - 1):
                    folder_readable_name = splitname[i] + '/'
                folder_readable_name = folder_readable_name + splitname[-1]
            else:
                folder_readable_name = result.groups()[1].split(seperator)[0]
            return folder_readable_name
        return False

    def _get_folders(self, cr, uid, context=None):
        if 'active_ids' in context.keys():
            record = self.pool.get(
                        'poweremail.core_accounts'
                        ).browse(cr, uid, context['active_ids'][0], context)
            if record:
                folderlist = []
                try:
                    if record.isssl:
                        serv = imaplib.IMAP4_SSL(record.iserver, record.isport)
                    else:
                        serv = imaplib.IMAP4(record.iserver, record.isport)
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Error"),
                                         _("An error occurred : %s ") % error)
                except Exception, error:
                    raise osv.except_osv(_("IMAP Server connection Error"),
                                         _("An error occurred : %s ") % error)
                try:
                    serv.login(record.isuser, record.ispass)
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Login Error"),
                                         _("An error occurred : %s ") % error)
                except Exception, error:
                    raise osv.except_osv(_("IMAP Server Login Error"),
                                         _("An error occurred : %s ") % error)
                try:
                    for folders in serv.list()[1]:
                        folder_readable_name = self.makereadable(folders)
                        if isinstance(folders, tuple):
                            data = folders[0] + folders[1]
                        else:
                            data = folders
                        if data.find('Noselect') == -1: #If it is a selectable folder
			    if folder_readable_name:
                                folderlist.append(
                                                  (folder_readable_name,
                                                   folder_readable_name)
                                                  )
                        if folder_readable_name == 'INBOX':
                            self.inboxvalue = folder_readable_name
                except imaplib.IMAP4.error, error:
                    raise osv.except_osv(_("IMAP Server Folder Error"),
                                         _("An error occurred : %s ") % error)
                except Exception, error:
                    raise osv.except_osv(_("IMAP Server Folder Error"),
                                         _("An error occurred : %s ") % error)
            else:
                folderlist = [('invalid', 'Invalid')]
        else:
            folderlist = [('invalid', 'Invalid')]
        return folderlist

    _columns = {
        'name':fields.many2one(
                        'poweremail.core_accounts',
                        string='Email Account',
                        readonly=True),
        'folder':fields.selection(
                        _get_folders,
                        string="IMAP Folder"),
    }

    _defaults = {
        'name':lambda self, cr, uid, ctx: ctx['active_ids'][0],
        'folder': lambda self, cr, uid, ctx:self.inboxvalue
    }

    def sel_folder(self, cr, uid, ids, context=None):
        """
        TODO: Doc This
        """
        if self.read(cr, uid, ids, ['folder'], context)[0]['folder']:
            if not self.read(cr, uid, ids,
                             ['folder'], context)[0]['folder'] == 'invalid':
                self.pool.get(
                        'poweremail.core_accounts'
                            ).write(cr, uid, context['active_ids'][0],
                                    {
                                     'isfolder':self.read(cr, uid, ids,
                                                  ['folder'],
                                                  context)[0]['folder']
                                    })
                return {
                        'type':'ir.actions.act_window_close'
                        }
            else:
                raise osv.except_osv(
                                     _("Folder Error"),
                                     _("This is an invalid folder "))
        else:
            raise osv.except_osv(
                                 _("Folder Error"),
                                 _("Select a folder before you save record "))
PoweremailSelectFolder()

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
