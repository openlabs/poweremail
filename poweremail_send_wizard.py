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

class poweremail_send_wizard(osv.osv_memory):
    _name = 'poweremail.send.wizard'
    _description = 'This is the wizard for sending mail'
    _rec_name = "subject"

    def _get_accounts(self,cr,uid,ctx={}):
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
                    accounts_id = self.pool.get('poweremail.core_accounts').search(cr,uid,[('company','=','yes'),('user','=',uid)])
                    if accounts_id:
                        accounts = self.pool.get('poweremail.core_accounts').browse(cr,uid,accounts_id)
                        return [(r.id,r.name + " (" + r.email_id + ")") for r in accounts]

    def get_value(self,cr,uid,ctx={},message={}):
        return self.engine.parsevalue(cr,uid,ctx['active_id'],message,self.template.id,ctx)
        

    _columns = {
        'ref_template':fields.many2one('poweremail.templates','Template',readonly=True),
        'rel_model':fields.many2one('ir.model','Model',readonly=True),
        #'rel_model_ref':fields.selection(_get_model_recs,'Referred Document',readonly=True),
        'from':fields.selection(_get_accounts,'From Account',required=True),
        'to':fields.char('To',size=100,readonly=True),
        'cc':fields.char('CC',size=100,),
        'bcc':fields.char('BCC',size=100,),
        'subject':fields.char('Subject',size=200,),
        'body_text':fields.text('Body',),
        'body_html':fields.text('Body',),
        'report':fields.char('Report File Name',size=100,),
        'signature':fields.boolean('Attach my signature to mail')
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
        #'signature': lambda self,cr,uid,ctx: self.template.use_sign
    }

poweremail_send_wizard()
    
    