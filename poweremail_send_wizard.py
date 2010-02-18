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
from mako import exceptions
import netsvc
import base64
from tools.translate import _
import tools

class poweremail_send_wizard(osv.osv_memory):
    _name = 'poweremail.send.wizard'
    _description = 'This is the wizard for sending mail'
    _rec_name = "subject"

    def _get_accounts(self,cr,uid,context=None):
        if context is None:
            context = {}
        logger = netsvc.Logger()
        tmpl_id = False
        if 'template' in context.keys():
            self.model_ref = context['src_rec_ids'][0]
            tmpl_id = self.pool.get('poweremail.templates').search(cr,uid,[('name','=',context['template'])], context=context)
        self.ctx = context
        if tmpl_id:
            self.template = self.pool.get('poweremail.templates').browse(cr,uid,tmpl_id[0], context)
            #Search translated template
            lang = self.get_value(cr,uid,context,self.template.lang)
            if lang:
                ctx = context.copy()
                ctx.update({'lang':lang})
                self.template = self.pool.get('poweremail.templates').browse(cr,uid,tmpl_id[0],ctx)
            #print self.template.allowed_groups
            if self.template.enforce_from_account:
                return [(self.template.enforce_from_account.id,self.template.enforce_from_account.name + " (" + self.template.enforce_from_account.email_id + ")")]
            else:
                accounts_id = self.pool.get('poweremail.core_accounts').search(cr,uid,[('company','=','no'),('user','=',uid)], context=context)
                if accounts_id:
                    accounts = self.pool.get('poweremail.core_accounts').browse(cr,uid,accounts_id, context)
                    return [(r.id,r.name + " (" + r.email_id + ")") for r in accounts]
                else:
                   logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("No personal email accounts are configured for you. \nEither ask admin to enforce an account for this template or get yourself a personal power email account."))
                   raise osv.except_osv(_("Power Email"),_("No personal email accounts are configured for you. \nEither ask admin to enforce an account for this template or get yourself a personal power email account."))

    def _get_generated(self,cr,uid,ids=None,context=None):
        if ids is None:
            ids = []
        if context is None:
            context = {}
        logger = netsvc.Logger()
        screen_vals = self.read(cr,uid,ids[0],[], context)
        context['account_id'] = screen_vals[0]['from']
        if self.ctx['src_rec_ids'] and len(self.ctx['src_rec_ids'])>1 and self.template.id:
            #Means there are multiple items selected for email. Just send them no need preview
            result =self.pool.get('poweremail.templates').generate_mail(cr,uid,self.template.id,self.ctx['src_rec_ids'],context) 
            if result:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Emails for multiple items saved in outbox."))
                self.write(cr,uid,ids,{'generated':len(result),'state':'done'}, context)
                #return {'value':{'generated':len(result)}}
            else:
                raise osv.except_osv(_("Power Email"),_("Email sending failed for one or more objects."))
        return True

    def get_value(self,cr,uid,context=None,message=None):
        if context is None:
            context = {}
        if message is None:
            message = {}
        if message:
            try:
                message = tools.ustr(message)
                object = self.pool.get(self.template.model_int_name).browse(cr,uid,ctx['src_rec_ids'][0],context)
                templ = Template(message,input_encoding='utf-8')
                env = {
                    'user':self.pool.get('res.users').browse(cr,uid,uid, context),
                    'db':cr.dbname
                       }
                reply = Template(message).render_unicode(object=object,peobject=object,env=env,format_exceptions=True)
                return reply
            except Exception:
                return ""
        else:
            return ""

    _columns = {
        'state':fields.selection([
                        ('single','Simple Mail Wizard Step 1'),
                        ('multi','Multiple Mail Wizard Step 1'),
                        ('done','Wizard Complete')
                                  ],'Status',readonly=True),
        'ref_template':fields.many2one('poweremail.templates','Template',readonly=True),
        'rel_model':fields.many2one('ir.model','Model',readonly=True),
        'rel_model_ref':fields.integer('Referred Document',readonly=True),
        'from':fields.selection(_get_accounts,'From Account',select=True),
        'to':fields.char('To',size=250,required=True),
        'cc':fields.char('CC',size=250,),
        'bcc':fields.char('BCC',size=250,),
        'subject':fields.char('Subject',size=200),
        'body_text':fields.text('Body',),
        'body_html':fields.text('Body',),
        'report':fields.char('Report File Name',size=100,),
        'signature':fields.boolean('Attach my signature to mail'),
        #'filename':fields.text('File Name'),
        'requested':fields.integer('No of requested Mails',readonly=True),
        'generated':fields.integer('No of generated Mails',readonly=True), 
        'full_success':fields.boolean('Complete Success',readonly=True)
                }

    _defaults = {
        'state': lambda self,cr,uid,ctx: len(ctx['src_rec_ids']) > 1 and 'multi' or 'single',
        'rel_model': lambda self,cr,uid,ctx:self.pool.get('ir.model').search(cr,uid,[('model','=',ctx['src_model'])],context=ctx)[0],
        'rel_model_ref': lambda self,cr,uid,ctx:ctx['active_id'],
        'to': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_to),
        'cc': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_cc),
        'bcc': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_bcc),
        'subject':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_subject),
        'body_text':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_body_text),
        'body_html':lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.def_body_html),
        'report': lambda self,cr,uid,ctx: self.get_value(cr,uid,ctx,self.template.file_name),
        'signature': lambda self,cr,uid,ctx: self.template.use_sign,
        'ref_template':lambda self,cr,uid,ctx: self.template.id,
        'requested':lambda self,cr,uid,ctx: len(ctx['src_rec_ids']),
        'full_success': lambda *a:False
    }

    def sav_to_drafts(self,cr,uid,ids,context=None):
        if context is None:
            context = {}
        mailid = self.save_to_mailbox(cr,uid,ids,context)
        if self.pool.get('poweremail.mailbox').write(cr,uid,mailid,{'folder':'drafts'}, context):
            return {'type':'ir.actions.act_window_close' }

    def send_mail(self,cr,uid,ids,context=None):
        if context is None:
            context = {}
        mailid = self.save_to_mailbox(cr,uid,ids,context)
        if self.pool.get('poweremail.mailbox').write(cr,uid,mailid,{'folder':'outbox'}, context):
            return {'type':'ir.actions.act_window_close' }
        
    def save_to_mailbox(self,cr,uid,ids,context=None):
        for id in ids:
            screen_vals = self.read(cr,uid,id,[], context)[0]
            accounts = self.pool.get('poweremail.core_accounts').read(cr,uid,screen_vals['from'], context)
            vals = {
                'pem_from': tools.ustr(accounts['name']) + "<" + tools.ustr(accounts['email_id']) + ">",
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
                sign = self.pool.get('res.users').read(cr,uid,uid,['signature'], context)['signature']
                if vals['pem_body_text']:
                    vals['pem_body_text'] = tools.ustr(vals['pem_body_text'])+sign
                if vals['pem_body_html']:
                    vals['pem_body_html'] = tools.ustr(vals['pem_body_html'])+sign
            #Create partly the mail and later update attachments
            mail_id = self.pool.get('poweremail.mailbox').create(cr,uid,vals, context)
            if self.template.report_template:
                record_id = screen_vals['rel_model_ref']
                reportname = 'report.' + self.pool.get('ir.actions.report.xml').read(cr,uid,self.template.report_template.id,['report_name'], context)['report_name']
                service = netsvc.LocalService(reportname)
                data = {}
                data['model'] = self.pool.get('ir.model').browse(cr, uid, screen_vals['rel_model'], context).model
                (result, format) = service.create(cr, uid, [record_id], data, context)
                att_obj = self.pool.get('ir.attachment')
                new_att_vals={
                                'name': _('%s (Email Attachment)') % screen_vals['subject'],
                                'datas':base64.b64encode(result),
                                'datas_fname':tools.ustr(screen_vals['report'] or _('Report')) + "." + format,
                                'description':screen_vals['body_text'] or _("No Description"),
                                'res_model':'poweremail.mailbox',
                                'res_id':mail_id
                                    }
                attid = att_obj.create(cr,uid,new_att_vals, context)
                if attid:
                    self.pool.get('poweremail.mailbox').write(cr,uid,mail_id,{'pem_attachments_ids':[[6, 0, [attid]]],'mail_type':'multipart/mixed'}, context)
            return mail_id
poweremail_send_wizard()

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
