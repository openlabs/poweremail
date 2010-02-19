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

        template = self._get_template(cr, uid, context)
        if not template:
            return []

        logger = netsvc.Logger()

        if template.enforce_from_account:
            return [(template.enforce_from_account.id, '%s (%s)' % (template.enforce_from_account.name, template.enforce_from_account.email_id))]
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
        template = self._get_template(cr, uid, context)
        if context['src_rec_ids'] and len(context['src_rec_ids'])>1 and template:
            #Means there are multiple items selected for email. Just send them no need preview
            result = self.pool.get('poweremail.templates').generate_mail(cr,uid,template.id,context['src_rec_ids'],context) 
            if result:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_INFO, _("Emails for multiple items saved in outbox."))
                self.write(cr,uid,ids,{
                    'generated':len(result),
                    'state':'done'
                }, context)
            else:
                raise osv.except_osv(_("Power Email"),_("Email sending failed for one or more objects."))
        return True

    def get_value(self,cr,uid, template, message, context=None):
        if not message:
            return ''
        try:
            message = tools.ustr(message)
            object = self.pool.get(template.model_int_name).browse(cr,uid,context['src_rec_ids'][0],context)
            templ = Template(message,input_encoding='utf-8')
            env = {
                'user':self.pool.get('res.users').browse(cr,uid,uid, context),
                'db':cr.dbname
            }
            reply = Template(message).render_unicode(object=object,peobject=object,env=env,format_exceptions=True)
            return reply
        except Exception, e:
            return ""

    def _get_template(self, cr, uid, context=None):
        if context is None:
            context = {}
        if not 'template' in context:
            return None
        if 'template' in context.keys():
            template_ids = self.pool.get('poweremail.templates').search(cr, uid, [('name','=',context['template'])], context=context)
        if not template_ids:
            return None

        template = self.pool.get('poweremail.templates').browse(cr, uid, template_ids[0], context)

        lang = self.get_value( cr, uid, template, template.lang, context )
        if lang:
            # Use translated template if necessary
            ctx = context.copy()
            ctx['lang'] = lang
            template = self.pool.get('poweremail.templates').browse(cr, uid, template.id, ctx)
        return template

    def _get_template_value(self, cr, uid, field, context=None):
        template = self._get_template(cr, uid, context)
        if not template:
            return False
        return self.get_value( cr, uid, template, getattr(template, field), context )

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
        'to': lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_to', ctx),
        'cc': lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_cc', ctx),
        'bcc': lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_bcc', ctx),
        'subject':lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_subject', ctx),
        'body_text':lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_body_text', ctx),
        'body_html':lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'def_body_html', ctx),
        'report': lambda self,cr,uid,ctx: self._get_template_value(cr, uid, 'file_name', ctx),
        'signature': lambda self,cr,uid,ctx: self._get_template(cr, uid, ctx).use_sign,
        'ref_template':lambda self,cr,uid,ctx: self._get_template(cr, uid, ctx).id,
        'requested':lambda self,cr,uid,ctx: len(ctx['src_rec_ids']),
        'full_success': lambda *a: False
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
            screen_vals = self.read(cr,uid,id,[],context)[0]
            accounts = self.pool.get('poweremail.core_accounts').read(cr, uid, screen_vals['from'], context=context)
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
                signature = self.pool.get('res.users').read(cr,uid,uid,['signature'], context)['signature']
                vals['pem_body_text'] = tools.ustr(vals['pem_body_text'] or '') + signature
                vals['pem_body_html'] = tools.ustr(vals['pem_body_html'] or '') + signature
            #Create partly the mail and later update attachments
            mail_id = self.pool.get('poweremail.mailbox').create(cr,uid,vals, context)
            template = self._get_template(cr, uid, context)
            if template.report_template:
                record_id = screen_vals['rel_model_ref']
                reportname = 'report.' + self.pool.get('ir.actions.report.xml').read(cr,uid,template.report_template.id,['report_name'], context)['report_name']
                data = {}
                data['model'] = self.pool.get('ir.model').browse(cr, uid, screen_vals['rel_model'], context).model

                # Ensure report is rendered using template's language
                ctx = context.copy()
                if template.lang:
                    ctx['lang'] = self.get_value( cr, uid, template, template.lang, context )
                service = netsvc.LocalService(reportname)
                (result, format) = service.create(cr, uid, [record_id], data, ctx)

                attachment_id = self.pool.get('ir.attachment').create(cr, uid, {
                    'name': _('%s (Email Attachment)') % screen_vals['subject'],
                    'datas': base64.b64encode(result),
                    'datas_fname': tools.ustr(screen_vals['report'] or _('Report')) + "." + format,
                    'description': screen_vals['body_text'] or _("No Description"),
                    'res_model': 'poweremail.mailbox',
                    'res_id': mail_id
                }, context)
                if attachment_id:
                    self.pool.get('poweremail.mailbox').write(cr, uid, mail_id, {
                        'pem_attachments_ids': [[6, 0, [attachment_id]]],
                        'mail_type':'multipart/mixed'
                    }, context)
            return mail_id
poweremail_send_wizard()

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
