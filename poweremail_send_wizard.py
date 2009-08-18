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
# any later version.                                                    #
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
from mako.template import Template
import netsvc
import base64

class poweremail_send_wizard(osv.osv_memory):
    _name = 'poweremail.send.wizard'
    _description = 'This is the wizard for sending mail'
    _rec_name = "subject"

    def _get_accounts(self,cr,uid,ctx={}):
        logger = netsvc.Logger()
        self.engine = self.pool.get("poweremail.engines")
        if 'template' in ctx.keys():
            self.model_ref = ctx['active_id']
            tmpl_id = self.pool.get('poweremail.templates').search(cr,uid,[('name','=',ctx['template'])])
            if tmpl_id:
                self.template = self.pool.get('poweremail.templates').browse(cr,uid,tmpl_id[0])
                #print self.template.allowed_groups
                if self.template.enforce_from_account:
                    return [(self.template.enforce_from_account.id,self.template.enforce_from_account.name + " (" + self.template.enforce_from_account.email_id + ")")]
                else:
                    accounts_id = self.pool.get('poweremail.core_accounts').search(cr,uid,[('company','=','no'),('user','=',uid)])
                    if accounts_id:
                        accounts = self.pool.get('poweremail.core_accounts').browse(cr,uid,accounts_id)
                        return [(r.id,r.name + " (" + r.email_id + ")") for r in accounts]
                    else:
                       logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("No personal email accounts are configured for you. \nEither ask admin to enforce an account for this template or get yourself a personal power email account."))

#    def get_value(self,cr,uid,ctx={},message={}):
#        if message:
#            return self.engine.parsevalue(cr,uid,ctx['active_id'],message,self.template.id,ctx)
#        else:
#            return ""

    def get_value(self,cr,uid,ctx={},message={}):
        if message:
            object = self.pool.get(self.template.model_int_name).browse(cr,uid,ctx['active_id'])
            reply = Template(message).render(object=object)
            print reply
            return reply
        else:
            return ""

    _columns = {
        'ref_template':fields.many2one('poweremail.templates','Template',readonly=True),
        'rel_model':fields.many2one('ir.model','Model',readonly=True),
        'rel_model_ref':fields.integer('Referred Document',readonly=True),
        'from':fields.selection(_get_accounts,'From Account',required=True),
        'to':fields.char('To',size=100,readonly=True),
        'cc':fields.char('CC',size=100,),
        'bcc':fields.char('BCC',size=100,),
        'subject':fields.char('Subject',size=200,required=True),
        'body_text':fields.text('Body',),
        'body_html':fields.text('Body',),
        'report':fields.char('Report File Name',size=100,),
        'signature':fields.boolean('Attach my signature to mail'),
        'filename':fields.text('File Name')
                }

    _defaults = {
        'rel_model': lambda self,cr,uid,ctx:self.pool.get('ir.model').search(cr,uid,[('model','=',ctx['src_model'])])[0],
        'rel_model_ref': lambda self,cr,uid,ctx:ctx['active_id'],
        'to': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_to),
        'cc': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_cc),
        'bcc': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_bcc),
        'subject':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_subject),
        'body_text':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_body_text),
        'body_html':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_body_html),
        'report': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.file_name),
        'signature': lambda self,cr,uid,ctx: self.template.use_sign,
        'ref_template':lambda self,cr,uid,ctx: self.template.id
    }

    def sav_to_drafts(self,cr,uid,ids,ctx={}):
        mailid = self.save_to_mailbox(cr,uid,ids,ctx)
        if self.pool.get('poweremail.mailbox').write(cr,uid,mailid,{'folder':'drafts'}):
            return {'type':'ir.actions.act_window_close' }

    def send_mail(self,cr,uid,ids,ctx={}):
        mailid = self.save_to_mailbox(cr,uid,ids,ctx)
        if self.pool.get('poweremail.mailbox').write(cr,uid,mailid,{'folder':'outbox'}):
            return {'type':'ir.actions.act_window_close' }
        
    def save_to_mailbox(self,cr,uid,ids,ctx={}):
        for id in ids:
            screen_vals = self.read(cr,uid,id,[])[0]
            print screen_vals
            accounts = self.pool.get('poweremail.core_accounts').read(cr,uid,screen_vals['from'])
            vals = {
                'pem_from': ctx['src_model'],
                'pem_to':screen_vals['to'],
                'pem_cc':screen_vals['cc'],
                'pem_bcc':screen_vals['bcc'],
                'pem_subject':screen_vals['subject'],
                'pem_body_text':screen_vals['body_text'],
                'pem_body_html':screen_vals['body_html'],
                'pem_account_id' :screen_vals['from'],
                'state':'na',
                'mail_type':'multipart/alternative' #Options:'multipart/mixed','multipart/alternative','text/plain','text/html'
            }
            if screen_vals['signature']:
                sign = self.pool.get('res.users').read(cr,uid,uid,['signature'])['signature']
                if vals['pem_body_text']:
                    vals['pem_body_text']+=sign
                if vals['pem_body_html']:
                    vals['pem_body_html']+=sign
            #Create partly the mail and later update attachments
            mail_id = self.pool.get('poweremail.mailbox').create(cr,uid,vals)
            if self.template.report_template:
                reportname = 'report.' + self.pool.get('ir.actions.report.xml').read(cr,uid,self.template.report_template.id,['report_name'])['report_name']
                service = netsvc.LocalService(reportname)
                (result, format) = service.create(cr, uid, [id], {}, {})
                att_obj = self.pool.get('ir.attachment')
                new_att_vals={
                                'name':screen_vals['subject'] + ' (Email Attachment)',
                                'datas':base64.b64encode(result),
                                'datas_fname':str(screen_vals['filename'] or 'Report') + "." + format,
                                'description':screen_vals['body_text'] or "No Description",
                                'res_model':'poweremail.mailbox',
                                'res_id':mail_id
                                    }
                attid = att_obj.create(cr,uid,new_att_vals)
                if attid:
                    self.pool.get('poweremail.mailbox').write(cr,uid,mail_id,{'pem_attachments_ids':[[6, 0, [attid]]],'mail_type':'multipart/mixed'})
            return mail_id
        
poweremail_send_wizard()
    
    