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
import poweremail_engines

class poweremail_templates(osv.osv):
    _name="poweremail.templates"
    _description = 'Power Email Templates for Models'

    def change_model(self,cr,uid,ids,object_name,ctx={}):
       mod_name = self.pool.get('ir.model').read(cr,uid,object_name,['model'])['model']
       return {'value':{'model_int_name':mod_name}}
       

    _columns = {
        'name' : fields.char('Name of Template',size=100,required=True),
        'object_name':fields.many2one('ir.model','Model'),
        'model_int_name':fields.char('Model Internal Name',size=200,),
        'def_to':fields.char('Recepient (To)',size=64,help="The default recepient of email. Placeholders can be used here."),
        'def_cc':fields.char('Default CC',size=64,help="The default CC for the email. Placeholders can be used here."),
        'def_bcc':fields.char('Default BCC',size=64,help="The default BCC for the email. Placeholders can be used here."),
        'def_subject':fields.char('Default Subject',size=200, help="The default subject of email. Placeholders can be used here."),
        'def_body_text':fields.text('Standard Body (Text)',help="The text version of the mail"),
        'def_body_html':fields.text('Body (Text-Web Client Only)',help="The text version of the mail"),
        'use_sign':fields.boolean('Use Signature',help="the signature from the User details will be appened to the mail"),
        'file_name':fields.char('File Name Pattern',size=200,help="File name pattern can be specified with placeholders. eg. 2009_SO003.pdf"),
        'report_template':fields.many2one('ir.actions.report.xml','Report to send'),
        #'report_template':fields.reference('Report to send',[('ir.actions.report.xml','Reports')],size=128),
        'allowed_groups':fields.many2many('res.groups','template_group_rel','templ_id','group_id',string="Allowed User Groups",  help="Only users from these groups will be allowed to send mails from this ID"),
        'enforce_from_account':fields.many2one('poweremail.core_accounts',string="Enforce From Account",help="Emails will be sent only from this account.",domain="[('company','=','yes')]"),

        'auto_email':fields.boolean('Auto Email', help="Selecting Auto Email will create a server action for you which automatically sends mail after a new record is created.\nNote:Auto email can be enabled only after saving template."),
        #Referred Stuff - Dont delete even if template is deleted
        'attached_wkf':fields.many2one('workflow','Workflow'),
        'attached_activity':fields.many2one('workflow.activity','Activity'),
        #Referred Stuff - Delete these if template are deleted or they will crash the system
        'server_action':fields.many2one('ir.actions.server','Related Server Action',help="Corresponding server action is here.",ondelete="cascade"),
        'ref_ir_act_window':fields.many2one('ir.actions.act_window','Window Action',readonly=True,ondelete="cascade"),
        'ref_ir_value':fields.many2one('ir.values','Wizard Button',readonly=True,ondelete="cascade"),
        #Expression Builder fields
        #Simple Fields
        'model_object_field':fields.many2one('ir.model.fields',string="Field",help="Select the field from the model you want to use.\nIf it is a relationship field you will be able to choose the nested values in the box below\n(Note:If there are no values make sure you have selected the correct model)",store=False),
        'sub_object':fields.many2one('ir.model','Sub-model',help='When a relation field is used this field will show you the type of field you have selected',store=False),
        'sub_model_object_field':fields.many2one('ir.model.fields','Sub Field',help='When you choose relationship fields this field will specify the sub value you can use.',store=False),
        'null_value':fields.char('Null Value',help="This Value is used if the field is empty",size=50,store=False),
        'copyvalue':fields.char('Expression',size=100,help="Copy and paste the value in the location you want to use a system value.",store=False),
        #Table Fields
        'table_model_object_field':fields.many2one('ir.model.fields',string="Table Field",help="Select the field from the model you want to use.\nOnly one2many & many2many fields can be used for tables)",store=False),
        'table_sub_object':fields.many2one('ir.model','Table-model',help='This field shows the model you will be using for your table',store=False),
        'table_required_fields':fields.many2many('ir.model.fields','fields_table_rel','field_id','table_id',string="Required Fields",help="Select the fieldsyou require in the table)",store=False),
        'table_html':fields.text('HTML code',help="Copy this html code to your HTML message body for displaying the info in your mail.",store=False)
    }

    _defaults = {

    }
    _sql_constraints = [
        ('name', 'unique (name)', 'The template name must be unique !')
    ]

    def create(self, cr, uid, vals, *args, **kwargs):
        src_obj = self.pool.get('ir.model').read(cr,uid,vals['object_name'],['model'])['model']
        win_val={
             'name': vals['name'] + " Mail Form",
             'type':'ir.actions.act_window',
             'res_model':'poweremail.send.wizard',
             'src_model': src_obj,
             'view_type': 'form',
             'context': "{'src_model':'" + src_obj + "','template':'" + vals['name'] + "','src_rec_id':active_id,'src_rec_ids':active_ids}",
             'view_mode':'form,tree',
             'view_id':self.pool.get('ir.ui.view').search(cr,uid,[('name','=','poweremail.send.wizard.form')])[0],
             'target': 'new',
             'auto_refresh':1
             }
        vals['ref_ir_act_window']= self.pool.get('ir.actions.act_window').create(cr, uid, win_val)
        value_vals={
             'name': 'Send Mail(' + vals['name'] + ")",
             'model': src_obj,
             'key2': 'client_action_multi',    
             'value': "ir.actions.act_window,"+ str(vals['ref_ir_act_window']),
             'object':True,
             }
        vals['ref_ir_value'] = self.pool.get('ir.values').create(cr, uid, value_vals)
        return super(poweremail_templates,self).create(cr, uid, vals, *args, **kwargs)   

    def write(self,cr,uid,ids,datas={},ctx={}):
        if 'auto_email' in datas.keys():#Has the auto email button toggled?
            if datas['auto_email']: #If auto email was enabled
                #Create Server Action
                vals = {
                        'state':'poweremail',
                        'poweremail_template':ids[0],
                        'name':self.pool.get('poweremail.templates').read(cr,uid,ids[0],['name'])['name'],
                        'condition':'True',
                        'model_id':self.read(cr,uid,ids[0],['object_name'])['object_name'][0]
                    }
                datas['server_action'] = self.pool.get('ir.actions.server').create(cr,uid,vals)
                #Attach Workflow to server action
                #Check if workflow activity also changed
                if 'attached_activity' in datas.keys():
                    #The workflow has changed, so cancel the previous one
                    ref_wf_act = self.read(cr, uid, ids[0], ['attached_activity'])['attached_activity']
                    if ref_wf_act:          #Delete existing reference
                        self.pool.get('workflow.activity').write(cr,uid,ref_wf_act[0],{'action_id':False})
                    #Now attach the server action to newly selected workflow activity
                    self.pool.get('workflow.activity').write(cr,uid,datas['attached_activity'],{'action_id':datas['server_action']})
            else: #Auto email got disabled, so delete all workflow attachments and prev server action
                ref_sr_act = self.read(cr, uid, ids[0], ['server_action'])['server_action']
                ref_wf_act = self.read(cr, uid, ids[0], ['attached_activity'])['attached_activity']
                if ref_sr_act:          #Delete server action
                    self.pool.get('ir.actions.server').unlink(cr,uid,ref_sr_act[0])
                if ref_wf_act:          #Delete Server action reference in workflow
                    self.pool.get('workflow.activity').write(cr,uid,ref_wf_act[0],{'action_id':False})

        #If only attached workflow stage changed?
        elif 'attached_activity' in datas.keys():
            #The workflow only has changed, so cancel the previous one and add server action to new one
            ref_wf_act = self.read(cr, uid, ids[0], ['attached_activity'])['attached_activity']
            if ref_wf_act:          #Delete existing reference
                self.pool.get('workflow.activity').write(cr,uid,ref_wf_act[0],{'action_id':False})
            #Now attach the server action to newly selected workflow activity
            ref_sr_act = self.read(cr, uid, ids[0], ['server_action'])['server_action']
            self.pool.get('workflow.activity').write(cr,uid,datas['attached_activity'],{'action_id':ref_sr_act[0]})

        return super(poweremail_templates,self).write(cr, uid,ids, datas, ctx)

    def unlink(self, cr, uid, ids, ctx={}):
        for id in ids:
            try:
                ref_ir_act_window = self.read(cr, uid, id, ['ref_ir_act_window'])['ref_ir_act_window']
                ref_ir_value = self.read(cr, uid, id, ['ref_ir_value'])['ref_ir_value']
                ref_sr_act = self.read(cr, uid, id, ['server_action'])['server_action']
                ref_wf_act = self.read(cr, uid, id, ['attached_activity'])['attached_activity']
                print ref_ir_act_window,ref_ir_value
                if ref_ir_act_window:   #Delete Wizard buttin
                    self.pool.get('ir.actions.act_window').unlink(cr,uid,ref_ir_act_window[0])
                if ref_ir_value:
                    self.pool.get('ir.values').unlink(cr,uid,ref_ir_value[0])
                if ref_sr_act:          #Delete server action
                    self.pool.get('ir.actions.server').unlink(cr,uid,ref_sr_act[0])
                if ref_wf_act:          #Delete Server action reference in workflow
                    self.pool.get('workflow.activity').write(cr,uid,ref_wf_act[0],{'action_id':False})
                super(poweremail_templates,self).unlink(cr,uid,id)
            except:
                raise osv.except_osv(_("Warning"),_("Deletion of Record failed"))
                return False
        return True
    
    def compute_pl(self,model_object_field,sub_model_object_field,null_value):
        #Configure for MAKO
        copy_val = ''
        if model_object_field:
            copy_val = "${object." + model_object_field
        if sub_model_object_field:
            copy_val += "." + sub_model_object_field
        if null_value:
            copy_val += " or '" + null_value + "'"
        if model_object_field:
            copy_val += "}"
        return copy_val 
            
    def _onchange_model_object_field(self,cr,uid,ids,model_object_field):
        if model_object_field:
            result={}
            field_obj = self.pool.get('ir.model.fields').browse(cr,uid,model_object_field)
            #Check if field is relational
            if field_obj.ttype in ['many2one','one2many','many2many']:
                res_ids=self.pool.get('ir.model').search(cr,uid,[('model','=',field_obj.relation)])
                #print res_ids[0]
                if res_ids:
                    result['sub_object'] = res_ids[0]
                    result['copyvalue'] = self.compute_pl(False,False,False)
                    result['sub_model_object_field'] = False
                    result['null_value'] = False
                    return {'value':result}
            else:
                #Its a simple field... just compute placeholder
                    result['sub_object'] = False
                    result['copyvalue'] = self.compute_pl(field_obj.name,False,False)
                    result['sub_model_object_field'] = False
                    result['null_value'] = False
                    return {'value':result}
            
    def _onchange_sub_model_object_field(self,cr,uid,ids,model_object_field,sub_model_object_field):
        if model_object_field and sub_model_object_field:
            result={}
            field_obj = self.pool.get('ir.model.fields').browse(cr,uid,model_object_field)
            if field_obj.ttype in ['many2one','one2many','many2many']:
                res_ids=self.pool.get('ir.model').search(cr,uid,[('model','=',field_obj.relation)])
                sub_field_obj = self.pool.get('ir.model.fields').browse(cr,uid,sub_model_object_field)
                #print res_ids[0]
                if res_ids:
                    result['sub_object'] = res_ids[0]
                    result['copyvalue'] = self.compute_pl(field_obj.name,sub_field_obj.name,False)
                    result['sub_model_object_field'] = sub_model_object_field
                    result['null_value'] = False
                    return {'value':result}
            else:
                #Its a simple field... just compute placeholder
                    result['sub_object'] = False
                    result['copyvalue'] = self.compute_pl(field_obj.name,False,False)
                    result['sub_model_object_field'] = False
                    result['null_value'] = False
                    return {'value':result}

    def _onchange_null_value(self,cr,uid,ids,model_object_field,sub_model_object_field,null_value):
        if model_object_field and null_value:
            result={}
            field_obj = self.pool.get('ir.model.fields').browse(cr,uid,model_object_field)
            if field_obj.ttype in ['many2one','one2many','many2many']:
                res_ids=self.pool.get('ir.model').search(cr,uid,[('model','=',field_obj.relation)])
                sub_field_obj = self.pool.get('ir.model.fields').browse(cr,uid,sub_model_object_field)
                #print res_ids[0]
                if res_ids:
                    result['sub_object'] = res_ids[0]
                    result['copyvalue'] = self.compute_pl(field_obj.name,sub_field_obj.name,null_value)
                    result['sub_model_object_field'] = sub_model_object_field
                    result['null_value'] = null_value
                    return {'value':result}
            else:
                #Its a simple field... just compute placeholder
                    result['sub_object'] = False
                    result['copyvalue'] = self.compute_pl(field_obj.name,False,null_value)
                    result['sub_model_object_field'] = False
                    result['null_value'] = null_value
                    return {'value':result}
                   
    def _onchange_table_model_object_field(self,cr,uid,ids,model_object_field):
        if model_object_field:
            result={}
            field_obj = self.pool.get('ir.model.fields').browse(cr,uid,model_object_field)
            if field_obj.ttype in ['many2one','one2many','many2many']:
                res_ids=self.pool.get('ir.model').search(cr,uid,[('model','=',field_obj.relation)])
                if res_ids:
                    result['table_sub_object'] = res_ids[0]
                    return {'value':result}
            else:
                #Its a simple field... just compute placeholder
                    result['sub_object'] = False
                    return {'value':result}
    
    def _onchange_table_required_fields(self,cr,uid,ids,table_model_object_field,table_required_fields):
        print table_model_object_field,table_required_fields
        if table_model_object_field and table_required_fields:
            result=''
            table_field_obj = self.pool.get('ir.model.fields').browse(cr,uid,table_model_object_field)
            field_obj = self.pool.get('ir.model.fields')         
            #Generate Html Header
            result +="<p>\n<table>\n<tr>"
            for each_rec in table_required_fields[0][2]:
                result += "\n<td>"
                record = field_obj.browse(cr,uid,each_rec)
                result += record.field_description
                result += "</td>"
            result +="\n</tr>\n"
            #Table header is defined,  now mako for table
            result += "%for o in object." + table_field_obj.name + ":\n<tr>"
            for each_rec in table_required_fields[0][2]:
                result += "\n<td>${o."
                record = field_obj.browse(cr,uid,each_rec)
                result += record.name
                result += "}</td>"
            result +="\n</tr>\n%endfor\n</table>\n</p>"
            return {'value':{'table_html':result}}

    def get_value(self,cr,uid,recid,message={},template=None):
        #Returns the computed expression
        if message:
            try:
                if not type(message) in [unicode]:
                    message = unicode(message,'UTF-8')
                object = self.pool.get(template.model_int_name).browse(cr,uid,recid)
                templ = Template(message,input_encoding='utf-8')
                reply = templ.render_unicode(object=object,peobject=object)
                return reply
            except Exception,e:
                return ""
        else:
            return ""
        
    def generate_mail(self,cr,uid,id,recids,context={}):
        #Generates an email an saves to outbox given the template id & record ID of a record in template's model
        #id: ID of template to be used
        #recid: record id for the mail
        #Context: 'account_id':The id of account to send from
        logger = netsvc.Logger()
        sent_recs = []
        from_account = False
        template = self.browse(cr,uid,id)
        #If account to send from is in context select it, else use enforced account 
        if 'account_id' in context.keys():
            from_account = self.pool.get('poweremail.core_accounts').read(cr,uid,context['account_id'],['name','email_id'])
        else:
            from_account = {'id':template.enforce_from_account.id, 'name':template.enforce_from_account.name, 'email_id':template.enforce_from_account.email_id}
        for recid in recids:
            try:
                self.engine = self.pool.get("poweremail.engines")
                
                vals = {
                        'pem_from': unicode(from_account['name']) + "<" + unicode(from_account['email_id']) + ">",
                        'pem_to':self.get_value(cr,uid,recid,template.def_to,template),
                        'pem_cc':self.get_value(cr,uid,recid,template.def_cc,template),
                        'pem_bcc':self.get_value(cr,uid,recid,template.def_bcc,template),
                        'pem_subject':self.get_value(cr,uid,recid,template.def_subject,template),
                        'pem_body_text':self.get_value(cr,uid,recid,template.def_body_text,template),
                        'pem_body_html':self.get_value(cr,uid,recid,template.def_body_html,template),
                        'pem_account_id' :from_account['id'],#This is a mandatory field when automatic emails are sent
                        'state':'na',
                        'folder':'outbox',
                        'mail_type':'multipart/alternative' #Options:'multipart/mixed','multipart/alternative','text/plain','text/html'
                    }
                if template.use_sign:
                    sign = self.pool.get('res.users').read(cr,uid,uid,['signature'])['signature']
                    if vals['pem_body_text']:
                        vals['pem_body_text']+=sign
                    if vals['pem_body_html']:
                        vals['pem_body_html']+=sign
                #Create partly the mail and later update attachments
                mail_id = self.pool.get('poweremail.mailbox').create(cr,uid,vals)
                if template.report_template:
                    reportname = 'report.' + self.pool.get('ir.actions.report.xml').read(cr,uid,template.report_template.id,['report_name'])['report_name']
                    service = netsvc.LocalService(reportname)
                    (result, format) = service.create(cr, uid, [recid], {}, {})
                    att_obj = self.pool.get('ir.attachment')
                    new_att_vals={
                                    'name':vals['pem_subject'] + ' (Email Attachment)',
                                    'datas':base64.b64encode(result),
                                    'datas_fname':unicode(self.get_value(cr,uid,recid,template.file_name,template) or 'Report') + "." + format,
                                    'description':vals['pem_body_text'] or "No Description",
                                    'res_model':'poweremail.mailbox',
                                    'res_id':mail_id
                                        }
                    attid = att_obj.create(cr,uid,new_att_vals)
                    if attid:
                        self.pool.get('poweremail.mailbox').write(cr,uid,mail_id,{'pem_attachments_ids':[[6, 0, [attid]]],'mail_type':'multipart/mixed'})
                sent_recs.append(recid)
            except Exception,error:
                logger.notifyChannel(_("Power Email"), netsvc.LOG_ERROR, _("Email Generation failed, Reason:%s")% (error))
                return sent_recs
        #all mails saved
        return sent_recs

poweremail_templates()

class poweremail_preview(osv.osv_memory):
    _name = "poweremail.preview"
    _description = "Power Email Template Preview"
    
    def _get_model_recs(self,cr,uid,ctx={}):
        #Fills up the selection box which allows records from the selected object to be displayed
        self.context = ctx
        if 'active_id' in ctx.keys():
            ref_obj_id = self.pool.get('poweremail.templates').read(cr,uid,ctx['active_id'],['object_name'])['object_name']
            ref_obj_name = self.pool.get('ir.model').read(cr,uid,ref_obj_id[0],['model'])['model']
            ref_obj_ids = self.pool.get(ref_obj_name).search(cr,uid,[])
            ref_obj_recs = self.pool.get(ref_obj_name).name_get(cr,uid,ref_obj_ids)
            #print ref_obj_recs
            return ref_obj_recs
    
    def get_value(self,cr,uid,recid,message={},template=None,ctx={}):
        #Returns the computed expression
        if message:
            try:
                if not type(message) in [unicode]:
                    message = unicode(message,'UTF-8')
                object = self.pool.get(template.model_int_name).browse(cr,uid,recid)
                reply = Template(message).render_unicode(object=object,peobject=object)
                return reply
            except Exception,e:
                return ""
        else:
            return ""
        
    _columns = {
        'ref_template':fields.many2one('poweremail.templates','Template',readonly=True),
        'rel_model':fields.many2one('ir.model','Model',readonly=True),
        'rel_model_ref':fields.selection(_get_model_recs,'Referred Document'),
        'to':fields.char('To',size=100,readonly=True),
        'cc':fields.char('CC',size=100,readonly=True),
        'bcc':fields.char('BCC',size=100,readonly=True),
        'subject':fields.char('Subject',size=200,readonly=True),
        'body_text':fields.text('Body',readonly=True),
        'body_html':fields.text('Body',readonly=True),
        'report':fields.char('Report Name',size=100,readonly=True),
    }
    _defaults = {
        'ref_template': lambda self,cr,uid,ctx:ctx['active_id'],
        'rel_model': lambda self,cr,uid,ctx:self.pool.get('poweremail.templates').read(cr,uid,ctx['active_id'],['object_name'])['object_name']
    }

    def _on_change_ref(self,cr,uid,ids,rel_model_ref,ctx={}):
        if rel_model_ref:
            vals={}
            if ctx == {}:
                ctx = self.context
            template = self.pool.get('poweremail.templates').browse(cr,uid,ctx['active_id'],ctx)
            vals['to']= self.get_value(cr,uid,rel_model_ref,template.def_to,template,ctx)
            vals['cc']= self.get_value(cr,uid,rel_model_ref,template.def_cc,template,ctx)
            vals['bcc']= self.get_value(cr,uid,rel_model_ref,template.def_bcc,template,ctx)
            vals['subject']= self.get_value(cr,uid,rel_model_ref,template.def_subject,template,ctx)
            vals['body_text']=self.get_value(cr,uid,rel_model_ref,template.def_body_text,template,ctx)
            vals['body_html']=self.get_value(cr,uid,rel_model_ref,template.def_body_html,template,ctx)
            vals['report']= self.get_value(cr,uid,rel_model_ref,template.file_name,template,ctx)
            #print "Vals>>>>>",vals
            return {'value':vals}
        
poweremail_preview()

class res_groups(osv.osv):
    _inherit = "res.groups"
    _description = "User Groups"
    _columns = {}
res_groups()