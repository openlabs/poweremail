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
import time
import poweremail_engines

class poweremail_mailbox(osv.osv):
    _name="poweremail.mailbox"
    _description = 'Power Email Mailbox included all type inbox,outbox,junk..'
    _rec_name="subject"

    def run_mail_scheduler(self, cr, uid, use_new_cursor=False, context=None):
        self.get_all_mail(cr,uid,context={'all_accounts':True})
        self.send_all_mail(cr,uid,context)
        
    def get_all_mail(self,cr,uid,context={}):
        #8888888888888 FETCHES MAILS 8888888888888888888#
        #email_account: THe ID of poweremil core account
        #Context should also have the last downloaded mail for an account
        #Normlly this function is expected to trigger from scheduler hence the value will not be there
        core_obj = self.pool.get('poweremail.core_accounts')
        if not 'all_accounts' in context.keys():
            #Get mails from that ID only
            core_obj.get_mails(cr,uid,[context['email_account']])
        else:
            accounts = core_obj.search(cr,uid,[('state','=','approved')])
            core_obj.get_mails(cr,uid,accounts)

    def get_fullmail(self,cr,uid,context={}):
        #8888888888888 FETCHES MAILS 8888888888888888888#
        core_obj = self.pool.get('poweremail.core_accounts')
        if 'mailboxref' in context.keys():
            #Get mails from that ID only
            core_obj.get_fullmail(cr,uid,context['email_account'],context)
        else:
            raise osv.osv_except(_("Mail fetch exception"),_("No information on which mail should be fetched fully"))
        
    def send_all_mail(self,cr,uid,ids=[],ctx={}):
        #8888888888888 SENDS MAILS IN OUTBOX 8888888888888888888#
        #get ids of mails in outbox
        filters = [('folder','=','outbox')]
        if 'filters' in ctx.keys():
            for each_filter in ctx['filters']:
                filters.append(each_filter)
        ids = self.search(cr,uid,filters)
        #send mails one by one
        for id in ids:
            core_obj=self.pool.get('poweremail.core_accounts')
            values =  self.read(cr,uid,id,[]) #Values will be a dictionary of all entries in the record ref by id
            payload={}
            if values['pem_attachments_ids']:
                #Get filenames & binary of attachments
                for attid in values['pem_attachments_ids']:
                    attachment = self.pool.get('ir.attachment').browse(cr,uid,attid)#,['datas_fname','datas'])
                    payload[attachment.datas_fname] = attachment.datas
            if core_obj.send_mail(cr,uid,[values['pem_account_id'][0]],[values['pem_to']or False],[values['pem_cc']or False],[values['pem_bcc']or False],values['pem_subject']or False,values['pem_body_text']or False,values['pem_body_html']or False,payload=payload):
                self.write(cr,uid,id,{'folder':'sent','state':'na','date_mail':time.strftime("%Y-%m-%d %H:%M:%S")})
    
    def send_this_mail(self,cr,uid,ids=[],ctx={}):
        #8888888888888 SENDS THIS MAIL IN OUTBOX 8888888888888888888#
        #send mails one by one
        for id in ids:
            core_obj=self.pool.get('poweremail.core_accounts')
            values =  self.read(cr,uid,id,[]) #Values will be a dictionary of all entries in the record ref by id
            payload={}
            if values['pem_attachments_ids']:
                #Get filenames & binary of attachments
                for attid in values['pem_attachments_ids']:
                    attachment = self.pool.get('ir.attachment').browse(cr,uid,attid)#,['datas_fname','datas'])
                    payload[attachment.datas_fname] = attachment.datas
            if core_obj.send_mail(cr,uid,[values['pem_account_id'][0]],[values['pem_to']or False],[values['pem_cc']or False],[values['pem_bcc']or False],values['pem_subject']or False,values['pem_body_text']or False,values['pem_body_html']or False,payload=payload):
                self.write(cr,uid,id,{'folder':'sent','state':'na','date_mail':time.strftime("%Y-%m-%d %H:%M:%S")})
                
    def complete_mail(self,cr,uid,ids,ctx={}):
        #8888888888888 COMPLETE PARTIALLY DOWNLOADED MAILS 8888888888888888888#
        #FUNCTION get_fullmail(self,cr,uid,mailid) in core is used where mailid=id of current email,
        for id in ids:
            self.pool.get('poweremail.core_accounts').get_fullmail(cr,uid,id,ctx)
    
    _columns = {
            'pem_from':fields.char('From', size=64),
            'pem_to':fields.char('Recepient (To)', size=64,),
            'pem_cc':fields.char(' CC', size=64),
            'pem_bcc':fields.char(' BCC', size=64),
            'pem_subject':fields.char(' Subject', size=200,),
            'pem_body_text':fields.text('Standard Body (Text)'),
            'pem_body_html':fields.text('Body (Text-Web Client Only)'),
            'pem_attachments_ids':fields.many2many('ir.attachment', 'mail_attachments_rel', 'mail_id', 'att_id', 'Attachments'),
            'pem_account_id' :fields.many2one('poweremail.core_accounts', 'User account'),
            'pem_user':fields.related('pem_account_id','user',type="many2one",relation="res.users",string="User"),
            'server_ref':fields.integer('Server Reference of mail',help="Applicable for inward items only"),
            'pem_recd':fields.char('Received at',size=50),
            'mail_type':fields.selection([
                                    ('multipart/mixed','Has Attachments'),
                                    ('multipart/alternative','Plain Text & HTML with no attachments'),
                                    ('multipart/related','Intermixed content'),
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
            'state':fields.selection([
                                    ('read','Read'),
                                    ('unread','Un-Read'),
                                    ('na','Not Applicable'),
                                    ],'Status'),
            'date_mail':fields.datetime('Rec/Sent Date')
        }

    _defaults = {

    } 



poweremail_mailbox()

    

