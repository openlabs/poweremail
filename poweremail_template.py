"""
Email templates & preview
"""
#########################################################################
#Power Email is a module for Open ERP which enables it to send mails    #
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
import base64
import random
import time
import types
import netsvc

LOGGER = netsvc.Logger()

TEMPLATE_ENGINES = []

from osv import osv, fields
from tools.translate import _
from tools.safe_eval import safe_eval
#Try and check the available templating engines
from mako.template import Template  #For backward combatibility
try:
    from mako.template import Template as MakoTemplate
    from mako import exceptions
    TEMPLATE_ENGINES.append(('mako', 'Mako Templates'))
except:
    LOGGER.notifyChannel(
                         _("Power Email"),
                         netsvc.LOG_ERROR,
                         _("Mako templates not installed")
                         )
try:
    from django.template import Context, Template as DjangoTemplate
    #Workaround for bug:
    #http://code.google.com/p/django-tagging/issues/detail?id=110
    from django.conf import settings
    settings.configure()
    #Workaround ends
    TEMPLATE_ENGINES.append(('django', 'Django Template'))
except:
    LOGGER.notifyChannel(
                         _("Power Email"),
                         netsvc.LOG_ERROR,
                         _("Django templates not installed")
                         )

import poweremail_engines
import tools
import report
import pooler

def send_on_create(self, cr, uid, vals, context=None):
    id = self.old_create(cr, uid, vals, context)
    template = self.pool.get('poweremail.templates').browse(cr, uid, self.template_id, context)
    # Ensure it's still configured to send on create
    if template.send_on_create:
        self.pool.get('poweremail.templates').generate_mail(cr, uid, self.template_id, [id], context)
    return id

def send_on_write(self, cr, uid, ids, vals, context=None):
    result = self.old_write(cr, uid, ids, vals, context)
    template = self.pool.get('poweremail.templates').browse(cr, uid, self.template_id, context)
    # Ensure it's still configured to send on write
    if template.send_on_write:
        self.pool.get('poweremail.templates').generate_mail(cr, uid, self.template_id, ids, context)
    return result


# This is an ugly hack to ensure that send_on_create and send_on_write are
# initialized when the server is started. Note there's a small time window
# between when the pool is available and when this function is called which
# may mean allow creating/writing objects without an e-mail being sent.


def new_register_all(db):
    cr = db.cursor()
    pool = pooler.get_pool(cr.dbname)
    value = pool.get('ir.actions.report.xml').register_all(cr)

    # If poweremail.templates has not yet been initialized, do not try to
    # SELECT its table yet
    if not 'poweremail.templates' in pool.obj_list():
        return value

    cr.execute("""
        SELECT
            pt.id,
            im.model,
            pt.send_on_create,
            pt.send_on_write
        FROM
            poweremail_templates pt,
            ir_model im
        WHERE
            pt.object_name = im.id
    """)
    for record in cr.fetchall():
        id = record[0]
        model = record[1]
        soc = record[2]
        sow = record[3]
        obj = pool.get(model)
        if not obj:
            continue
        if hasattr(obj, 'old_create'):
            obj.create = obj.old_create
            del obj.old_create
        if hasattr(obj, 'old_write'):
            obj.write = obj.old_write
            del obj.old_write
        if soc:
            obj.template_id = id
            obj.old_create = obj.create
            obj.create = types.MethodType(send_on_create, obj, osv.osv)
        if sow:
            obj.template_id = id
            obj.old_write = obj.write
            obj.write = types.MethodType(send_on_write, obj, osv.osv)

    cr.close()
    return value

report.interface.register_all = new_register_all

def get_value(cursor, user, recid, message=None, template=None, context=None):
    """
    Evaluates an expression and returns its value
    @param cursor: Database Cursor
    @param user: ID of current user
    @param recid: ID of the target record under evaluation
    @param message: The expression to be evaluated
    @param template: BrowseRecord object of the current template
    @param context: Open ERP Context
    @return: Computed message (unicode) or u""
    """
    pool = pooler.get_pool(cursor.dbname)
    if message is None:
        message = {}
    #Returns the computed expression
    if message:
        try:
            message = tools.ustr(message)
            object = pool.get(template.model_int_name).browse(cursor, user, recid, context)
            env = {
                'user':pool.get('res.users').browse(cursor, user, user, context),
                'db':cursor.dbname
                   }
            if template.template_language == 'mako':
                templ = MakoTemplate(message, input_encoding='utf-8')
                reply = MakoTemplate(message).render_unicode(object=object,
                                                             peobject=object,
                                                             env=env,
                                                             format_exceptions=True)
            elif template.template_language == 'django':
                templ = DjangoTemplate(message)
                env['object'] = object
                env['peobject'] = object
                reply = templ.render(Context(env))
            return reply or False
        except Exception:
            return u""
    else:
        return message

class poweremail_templates(osv.osv):
    "Templates for sending Email"

    _name = "poweremail.templates"
    _description = 'Power Email Templates for Models'

    def change_model(self, cursor, user, ids, object_name, context=None):
        if object_name:
            mod_name = self.pool.get('ir.model').read(
                                              cursor,
                                              user,
                                              object_name,
                                              ['model'], context)['model']
        else:
            mod_name = False
        return {
                'value':{'model_int_name':mod_name}
                }

    _columns = {
        'name' : fields.char('Name of Template', size=100, required=True),
        'object_name':fields.many2one('ir.model', 'Model'),
        'model_int_name':fields.char('Model Internal Name', size=200,),
        'def_to':fields.char(
                'Recepient (To)',
                size=250,
                help="The default recepient of email. "
                "Placeholders can be used here."),
        'def_cc':fields.char(
                'Default CC',
                size=250,
                help="The default CC for the email. "
                "Placeholders can be used here."),
        'def_bcc':fields.char(
                'Default BCC',
                size=250,
                help="The default BCC for the email. "
                "Placeholders can be used here."),
        'lang':fields.char(
                'Language',
                size=250,
                help="The default language for the email. "
                "Placeholders can be used here. "
                "eg. ${object.partner_id.lang}"),
        'def_subject':fields.char(
                'Default Subject',
                size=200,
                help="The default subject of email. "
                "Placeholders can be used here.",
                translate=True),
        'def_body_text':fields.text(
                'Standard Body (Text)',
                help="The text version of the mail.",
                translate=True),
        'def_body_html':fields.text(
                'Body (Text-Web Client Only)',
                help="The text version of the mail.",
                translate=True),
        'use_sign':fields.boolean(
                'Use Signature',
                help="The signature from the User details "
                "will be appened to the mail."),
        'file_name':fields.char(
                'File Name Pattern',
                size=200,
                help="File name pattern can be specified with placeholders. "
                "eg. 2009_SO003.pdf",
                translate=True),
        'report_template':fields.many2one(
                'ir.actions.report.xml',
                'Report to send'),
        #'report_template':fields.reference('Report to send',[('ir.actions.report.xml','Reports')],size=128),
        'allowed_groups':fields.many2many(
                'res.groups',
                'template_group_rel',
                'templ_id', 'group_id',
                string="Allowed User Groups",
                help="Only users from these groups will be "
                "allowed to send mails from this Template."),
        'enforce_from_account':fields.many2one(
                'poweremail.core_accounts',
                string="Enforce From Account",
                help="Emails will be sent only from this account.",
                domain="[('company','=','yes')]"),
        'auto_email':fields.boolean('Auto Email',
                help="Selecting Auto Email will create a server "
                "action for you which automatically sends mail after a "
                "new record is created.\nNote: Auto email can be enabled "
                "only after saving template."),
        'save_to_drafts':fields.boolean('Save to Drafts',
                    help="When automatically sending emails generated from"
                    " this template, save them into the Drafts folder rather"
                    " than sending them immediately."),
        #Referred Stuff - Dont delete even if template is deleted
        'attached_wkf':fields.many2one(
                'workflow',
                'Workflow'),
        'attached_activity':fields.many2one(
                'workflow.activity',
                'Activity'),
        #Referred Stuff - Delete these if template are deleted or they will crash the system
        'server_action':fields.many2one(
                'ir.actions.server',
                'Related Server Action',
                help="Corresponding server action is here."),
        'ref_ir_act_window':fields.many2one(
                'ir.actions.act_window',
                'Window Action',
                readonly=True),
        'ref_ir_value':fields.many2one(
                'ir.values',
                'Wizard Button',
               readonly=True),
        #Expression Builder fields
        #Simple Fields
        'model_object_field':fields.many2one(
                'ir.model.fields',
                string="Field",
                help="Select the field from the model you want to use."
                "\nIf it is a relationship field you will be able to "
                "choose the nested values in the box below.\n(Note: If "
                "there are no values make sure you have selected the "
                "correct model).",
                store=False),
        'sub_object':fields.many2one(
                'ir.model',
                'Sub-model',
                help='When a relation field is used this field '
                'will show you the type of field you have selected.',
                store=False),
        'sub_model_object_field':fields.many2one(
                'ir.model.fields',
                'Sub Field',
                help="When you choose relationship fields "
                "this field will specify the sub value you can use.",
                store=False),
        'null_value':fields.char(
                'Null Value',
                help="This Value is used if the field is empty.",
                size=50, store=False),
        'copyvalue':fields.char(
                'Expression',
                size=100,
                help="Copy and paste the value in the "
                "location you want to use a system value.",
                store=False),
        #Table Fields
        'table_model_object_field':fields.many2one(
                'ir.model.fields',
                string="Table Field",
                help="Select the field from the model you want to use."
                "\nOnly one2many & many2many fields can be used for tables.",
                store=False),
        'table_sub_object':fields.many2one(
                'ir.model',
                'Table-model',
                help="This field shows the model you will "
                "be using for your table.", store=False),
        'table_required_fields':fields.many2many(
                'ir.model.fields',
                'fields_table_rel',
                'field_id', 'table_id',
                string="Required Fields",
                help="Select the fields you require in the table.",
                store=False),
        'table_html':fields.text(
                'HTML code',
                help="Copy this html code to your HTML message "
                "body for displaying the info in your mail.",
                store=False),
        'send_on_create': fields.boolean(
                'Send on Create',
                help='Sends an e-mail when a new document is created.'),
        'send_on_write': fields.boolean(
                'Send on Update',
                help='Sends an e-mail when a document is modified.'),
        'partner_event': fields.char(
                'Partner ID to log Events',
                size=250,
                help="Partner ID who an email event is logged.\n"
                "Placeholders can be used here. eg. ${object.partner_id.id}\n"
                "You must install the mail_gateway module to see the mail events "
                "in partner form.\nIf you also want to record the link to the "
                "object that sends the email, you must to add this object in the "
                "'Administration/Low Level Objects/Requests/Accepted Links in "
                "Requests' menu (or 'ir.attachment' to record the attachments)."),
        'template_language':fields.selection(
                TEMPLATE_ENGINES,
                'Templating Language',
                required=True
                ),
                'single_email': fields.boolean("Single email", help="Check it if you want to send a single email for several records (the optional attachment will be generated as a single file for all these records). If you don't check it, an email with its optional attachment will be send for each record."),
        'use_filter':fields.boolean(
                    'Active Filter',
                    help="This option allow you to add a custom python filter"
                    " before sending a mail"),
        'filter':fields.text(
                    'Filter',
                    help="The python code entered here will be excecuted if the"
                    "result is True the mail will be send if it false the mail "
                    "won't be send.\n"
                    "Example : o.type == 'out_invoice' and o.number and o.number[:3]<>'os_' "),
    }

    _defaults = {

    }
    _sql_constraints = [
        ('name', 'unique (name)', _('The template name must be unique!'))
    ]

    def update_auto_email(self, cr, uid, ids, context=None):
        for template in self.browse(cr, uid, ids, context):
            if template.auto_email:
                if not template.server_action:
                    # Create server action if necessary
                    action_id = self.pool.get('ir.actions.server').create(cr, uid, {
                        'state': 'poweremail',
                        'poweremail_template': template.id,
                        'name': template.name,
                        'condition': 'True',
                        'model_id': template.object_name.id,
                    }, context)
                    self.write(cr, uid, template.id, {
                        'server_action': action_id,
                    }, context)
                    self.pool.get('workflow.activity').write(cr, uid, template.attached_activity.id, {
                        'action_id': action_id,
                    }, context)
                else:
                    # Update activity if it was changed
                    activity_ids = self.pool.get('workflow.activity').search(cr, uid, [('action_id', '=', template.server_action.id)], context=context)
                    if not template.attached_activity.id in activity_ids:
                        self.pool.get('workflow.activity').write(cr, uid, activity_ids, {
                            'action_id': False,
                        }, context)
                        if template.attached_activity.id:
                            self.pool.get('workflow.activity').write(cr, uid, template.attached_activity.id, {
                                'action_id': template.server_action.id,
                            }, context)
            elif template.server_action:
                    self.pool.get('ir.actions.server').unlink(cr, uid, template.server_action.id, context)

    def update_send_on_store(self, cr, uid, ids, context):
        for template in self.browse(cr, uid, ids, context):
            obj = self.pool.get(template.object_name.model)
            if hasattr(obj, 'old_create'):
                obj.create = obj.old_create
                del obj.old_create
            if hasattr(obj, 'old_write'):
                obj.write = obj.old_write
                del obj.old_write
            if template.send_on_create:
                obj.template_id = template.id
                obj.old_create = obj.create
                obj.create = types.MethodType(send_on_create, obj, osv.osv)
            if template.send_on_write:
                obj.template_id = template.id
                obj.old_write = obj.write
                obj.write = types.MethodType(send_on_write, obj, osv.osv)

    def create(self, cr, uid, vals, context=None):
        id = super(poweremail_templates, self).create(cr, uid, vals, context)
        src_obj = self.pool.get('ir.model').read(cr, uid, vals['object_name'], ['model'], context)['model']
        vals['ref_ir_act_window'] = self.pool.get('ir.actions.act_window').create(cr, uid, {
             'name': _("%s Mail Form") % vals['name'],
             'type': 'ir.actions.act_window',
             'res_model': 'poweremail.send.wizard',
             'src_model': src_obj,
             'view_type': 'form',
             'context': "{'src_model':'%s','template_id':'%d','src_rec_id':active_id,'src_rec_ids':active_ids}" % (src_obj, id),
             'view_mode':'form,tree',
             'view_id': self.pool.get('ir.ui.view').search(cr, uid, [('name', '=', 'poweremail.send.wizard.form')], context=context)[0],
             'target': 'new',
             'auto_refresh':1
        }, context)
        vals['ref_ir_value'] = self.pool.get('ir.values').create(cr, uid, {
             'name': _('Send Mail (%s)') % vals['name'],
             'model': src_obj,
             'key2': 'client_action_multi',
             'value': "ir.actions.act_window," + str(vals['ref_ir_act_window']),
             'object': True,
         }, context)
        self.write(cr, uid, id, {
            'ref_ir_act_window': vals['ref_ir_act_window'],
            'ref_ir_value': vals['ref_ir_value'],
        }, context)
        if vals.get('auto_email'):
            self.update_auto_email(cr, uid, [id], context)
        if vals.get('send_on_create') or vals.get('send_on_write'):
            self.update_send_on_store(cr, uid, [id], context)
        #if vals.get('partner_event'):
        #    self.update_partner_event(cr, uid, [id], context)
        return id

    def write(self, cr, uid, ids, vals, context=None):
        result = super(poweremail_templates, self).write(cr, uid, ids, vals, context)
        if 'auto_email' in vals or 'attached_activity' in vals:
            self.update_auto_email(cr, uid, ids, context)
        if 'send_on_create' in vals or 'send_on_write' in vals:
            self.update_send_on_store(cr, uid, ids, context)
        #if 'partner_event' in vals:
        #    self.update_partner_event(cr, uid, ids, context)
        return result

    def unlink(self, cr, uid, ids, context=None):
        for template in self.browse(cr, uid, ids, context):
            obj = self.pool.get(template.object_name.model)
            if hasattr(obj, 'old_create'):
                obj.create = obj.old_create
                del obj.old_create
            if hasattr(obj, 'old_write'):
                obj.write = obj.old_write
                del obj.old_write
            try:
                if template.ref_ir_act_window:
                    self.pool.get('ir.actions.act_window').unlink(cr, uid, template.ref_ir_act_window.id, context)
                if template.ref_ir_value:
                    self.pool.get('ir.values').unlink(cr, uid, template.ref_ir_value.id, context)
                if template.server_action:
                    self.pool.get('ir.actions.server').unlink(cr, uid, template.server_action.id, context)
            except:
                raise osv.except_osv(_("Warning"), _("Deletion of Record failed"))
        return super(poweremail_templates, self).unlink(cr, uid, ids, context)

    def copy(self, cr, uid, id, default=None, context=None):
        if default is None:
            default = {}
        default = default.copy()
        old = self.read(cr, uid, id, ['name'], context=context)
        new_name = _("Copy of template ") + old.get('name', 'No Name')
        check = self.search(cr, uid, [('name', '=', new_name)], context=context)
        if check:
            new_name = new_name + '_' + random.choice('abcdefghij') + random.choice('lmnopqrs') + random.choice('tuvwzyz')
        default.update({'name':new_name})
        return super(poweremail_templates, self).copy(cr, uid, id, default, context)

    def compute_pl(self,
                   model_object_field,
                   sub_model_object_field,
                   null_value, template_language='mako'):
        """
        Returns the expression based on data provided
        @param model_object_field: First level field
        @param sub_model_object_field: Second level drilled down field (M2O)
        @param null_value: What has to be returned if the value is empty
        @param template_language: The language used for templating
        @return: computed expression
        """
        #Configure for MAKO
        copy_val = ''
        if template_language == 'mako':
            if model_object_field:
                copy_val = "${object." + model_object_field
            if sub_model_object_field:
                copy_val += "." + sub_model_object_field
            if null_value:
                copy_val += " or '" + null_value + "'"
            if model_object_field:
                copy_val += "}"
        elif template_language == 'django':
            if model_object_field:
                copy_val = "{{object." + model_object_field
            if sub_model_object_field:
                copy_val += "." + sub_model_object_field
            if null_value:
                copy_val = copy_val + '|default:"' + null_value + '"'
            copy_val = copy_val + "}}"
        return copy_val

    def onchange_model_object_field(self, cr, uid, ids, model_object_field, template_language, context=None):
        if not model_object_field:
            return {}
        result = {}
        field_obj = self.pool.get('ir.model.fields').browse(cr, uid, model_object_field, context)
        #Check if field is relational
        if field_obj.ttype in ['many2one', 'one2many', 'many2many']:
            res_ids = self.pool.get('ir.model').search(cr, uid, [('model', '=', field_obj.relation)], context=context)
            if res_ids:
                result['sub_object'] = res_ids[0]
                result['copyvalue'] = self.compute_pl(False,
                                                      False,
                                                      False,
                                                      template_language)
                result['sub_model_object_field'] = False
                result['null_value'] = False
        else:
            #Its a simple field... just compute placeholder
            result['sub_object'] = False
            result['copyvalue'] = self.compute_pl(field_obj.name,
                                                  False,
                                                  False,
                                                  template_language
                                                  )
            result['sub_model_object_field'] = False
            result['null_value'] = False
        return {'value':result}

    def onchange_sub_model_object_field(self, cr, uid, ids, model_object_field, sub_model_object_field, template_language, context=None):
        if not model_object_field or not sub_model_object_field:
            return {}
        result = {}
        field_obj = self.pool.get('ir.model.fields').browse(cr, uid, model_object_field, context)
        if field_obj.ttype in ['many2one', 'one2many', 'many2many']:
            res_ids = self.pool.get('ir.model').search(cr, uid, [('model', '=', field_obj.relation)], context=context)
            sub_field_obj = self.pool.get('ir.model.fields').browse(cr, uid, sub_model_object_field, context)
            if res_ids:
                result['sub_object'] = res_ids[0]
                result['copyvalue'] = self.compute_pl(field_obj.name,
                                                      sub_field_obj.name,
                                                      False,
                                                      template_language
                                                      )
                result['sub_model_object_field'] = sub_model_object_field
                result['null_value'] = False
        else:
            #Its a simple field... just compute placeholder
            result['sub_object'] = False
            result['copyvalue'] = self.compute_pl(field_obj.name,
                                                  False,
                                                  False,
                                                  template_language
                                                  )
            result['sub_model_object_field'] = False
            result['null_value'] = False
        return {'value':result}

    def onchange_null_value(self, cr, uid, ids, model_object_field, sub_model_object_field, null_value, template_language, context=None):
        if not model_object_field and not null_value:
            return {}
        result = {}
        field_obj = self.pool.get('ir.model.fields').browse(cr, uid, model_object_field, context)
        if field_obj.ttype in ['many2one', 'one2many', 'many2many']:
            res_ids = self.pool.get('ir.model').search(cr, uid, [('model', '=', field_obj.relation)], context=context)
            sub_field_obj = self.pool.get('ir.model.fields').browse(cr, uid, sub_model_object_field, context)
            if res_ids:
                result['sub_object'] = res_ids[0]
                result['copyvalue'] = self.compute_pl(field_obj.name,
                                                      sub_field_obj.name,
                                                      null_value,
                                                      template_language
                                                      )
                result['sub_model_object_field'] = sub_model_object_field
                result['null_value'] = null_value
        else:
            #Its a simple field... just compute placeholder
            result['sub_object'] = False
            result['copyvalue'] = self.compute_pl(field_obj.name,
                                                  False,
                                                  null_value,
                                                  template_language
                                                  )
            result['sub_model_object_field'] = False
            result['null_value'] = null_value
        return {'value':result}

    def onchange_table_model_object_field(self, cr, uid, ids, model_object_field, template_language, context=None):
        if not model_object_field:
            return {}
        result = {}
        field_obj = self.pool.get('ir.model.fields').browse(cr, uid, model_object_field, context)
        if field_obj.ttype in ['many2one', 'one2many', 'many2many']:
            res_ids = self.pool.get('ir.model').search(cr, uid, [('model', '=', field_obj.relation)], context=context)
            if res_ids:
                result['table_sub_object'] = res_ids[0]
        else:
            #Its a simple field... just compute placeholder
            result['sub_object'] = False
        return {'value':result}

    def onchange_table_required_fields(self, cr, uid, ids, table_model_object_field, table_required_fields, template_language, context=None):
        if not table_model_object_field or not table_required_fields:
            return {'value':{'table_html': False}}
        result = ''
        table_field_obj = self.pool.get('ir.model.fields').browse(cr, uid, table_model_object_field, context)
        field_obj = self.pool.get('ir.model.fields')
        #Generate Html Header
        result += "<p>\n<table border='1'>\n<thead>\n<tr>"
        for each_rec in table_required_fields[0][2]:
            result += "\n<td>"
            record = field_obj.browse(cr, uid, each_rec, context)
            result += record.field_description
            result += "</td>"
        result += "\n</tr>\n</thead>\n<tbody>\n"
        #Table header is defined,  now mako for table
        #print "Language:", template_language
        if template_language == 'mako':
            result += "%for o in object." + table_field_obj.name + ":\n<tr>"
            for each_rec in table_required_fields[0][2]:
                result += "\n<td>${o."
                record = field_obj.browse(cr, uid, each_rec, context)
                result += record.name
                result += "}</td>"
            result += "\n</tr>\n%endfor\n</tbody>\n</table>\n</p>"
        elif template_language == 'django':
            result += "{% for o in object." + table_field_obj.name + " %}\n<tr>"
            for each_rec in table_required_fields[0][2]:
                result += "\n<td>{{o."
                record = field_obj.browse(cr, uid, each_rec, context)
                result += record.name
                result += "}}</td>"
            result += "\n</tr>\n{% endfor %}\n</tbody>\n</table>\n</p>"
        return {'value':{'table_html':result}}

    def _generate_partner_events(self,
                                 cursor,
                                 user,
                                 template,
                                 record_id,
                                 mail,
                                 context=None):
        """
        Generates partner event if configured
        @author: Jordi Esteve
        
        @param cursor: Database Cursor
        @param user: ID of User
        @param template: Browse record of
                         template
        @param record_id: ID of the target model
                          for which this mail has
                          to be generated
        @param mail: Browse record of email object
        
        @return: True
        """
        name = mail.pem_subject
        if isinstance(name, str):
            name = unicode(name, 'utf-8')
        if len(name) > 64:
            name = name[:61] + '...'
        model = res_id = False
        if template.report_template:
            if self.pool.get('res.request.link').search(
                                    cursor,
                                    user,
                                    [('object', '=', template.model_int_name)],
                                    context=context):
                model = template.model_int_name
                res_id = record_id
            elif mail.pem_attachments_ids \
                    and self.pool.get('res.request.link').search(
                                        cursor,
                                        user,
                                        [('object', '=', 'ir.attachment')],
                                        context=context):
                model = 'ir.attachment'
                res_id = mail.pem_attachments_ids[0]
        event_vals = {
            'history': True,
            'name': name,
            'date': time.strftime('%Y-%m-%d %H:%M:%S'),
            'user_id': user,
            'email_from': mail.pem_from or None,
            'email_to': mail.pem_to or None,
            'email_cc': mail.pem_cc or None,
            'email_bcc': mail.pem_bcc or None,
            'message_id': mail.id,
            'description': mail.pem_body_text and mail.pem_body_text or mail.pem_body_html,
            'partner_id': get_value(cursor, user, record_id, template.partner_event, template, context),
            'model': model,
            'res_id': res_id,
        }
        self.pool.get('mailgate.message').create(cursor,
                                                  user,
                                                  event_vals,
                                                  context)
        return True

    def _generate_attach_reports(self,
                                 cursor,
                                 user,
                                 template,
                                 record_ids,
                                 mail,
                                 context=None):
        """
        Generate report to be attached and attach it
        to the email
        
        @param cursor: Database Cursor
        @param user: ID of User
        @param template: Browse record of
                         template
        @param record_ids: IDs of the target model
                           for which this mail has
                           to be generated
        @param mail: Browse record of email object
        @return: True
        """
        lang = get_value(cursor,
                         user,
                         record_ids[0],
                         template.lang,
                         template,
                         context)
        if lang:
            ctx = context.copy()
            ctx.update({'lang':lang})
            template = self.browse(cursor, user, template.id, context=ctx)
        reportname = 'report.' + \
            self.pool.get('ir.actions.report.xml').read(
                                         cursor,
                                         user,
                                         template.report_template.id,
                                         ['report_name'],
                                         context)['report_name']
        service = netsvc.LocalService(reportname)
        data = {}
        data['model'] = template.model_int_name
        (result, format) = service.create(cursor,
                                          user,
                                          record_ids,
                                          data,
                                          context)
        attachment_obj = self.pool.get('ir.attachment')
        new_att_vals = {
            'name': mail.pem_subject + ' (Email Attachment)',
            'datas': base64.b64encode(result),
            'datas_fname': tools.ustr(
                             get_value(
                                   cursor,
                                   user,
                                   record_ids[0],
                                   template.file_name,
                                   template,
                                   context
                                   ) or 'Report') + "." + format,
            'description': mail.pem_subject or "No Description",
            'res_model': 'poweremail.mailbox',
            'res_id': mail.id
        }
        attachment_id = attachment_obj.create(cursor,
                                              user,
                                              new_att_vals,
                                              context)
        if attachment_id:
            self.pool.get('poweremail.mailbox').write(
                              cursor,
                              user,
                              mail.id,
                              {
                               'pem_attachments_ids':[
                                                  [6, 0, [attachment_id]]
                                                    ],
                               'mail_type':'multipart/mixed'
                               },
                               context)
        return True

    def _generate_mailbox_item_from_template(self,
                                      cursor,
                                      user,
                                      template,
                                      record_id,
                                      context=None):
        """
        Generates an email from the template for
        record record_id of target object
        
        @param cursor: Database Cursor
        @param user: ID of User
        @param template: Browse record of
                         template
        @param record_id: ID of the target model
                          for which this mail has
                          to be generated
        @return: ID of created object
        """
        if context is None:
            context = {}
        #If account to send from is in context select it, else use enforced account
        if 'account_id' in context.keys():
            from_account = self.pool.get('poweremail.core_accounts').read(
                                                    cursor,
                                                    user,
                                                    context.get('account_id'),
                                                    ['name', 'email_id'],
                                                    context
                                                    )
        else:
            from_account = {
                            'id':template.enforce_from_account.id,
                            'name':template.enforce_from_account.name,
                            'email_id':template.enforce_from_account.email_id
                            }
        lang = get_value(cursor,
                         user,
                         record_id,
                         template.lang,
                         template,
                         context)
        if lang:
            ctx = context.copy()
            ctx.update({'lang':lang})
            template = self.browse(cursor, user, template.id, context=ctx)
        mailbox_values = {
            'pem_from': tools.ustr(from_account['name']) + \
                        "<" + tools.ustr(from_account['email_id']) + ">",
            'pem_to':get_value(cursor,
                               user,
                               record_id,
                               template.def_to,
                               template,
                               context),
            'pem_cc':get_value(cursor,
                               user,
                               record_id,
                               template.def_cc,
                               template,
                               context),
            'pem_bcc':get_value(cursor,
                                user,
                                record_id,
                                template.def_bcc,
                                template,
                                context),
            'pem_subject':get_value(cursor,
                                    user,
                                    record_id,
                                    template.def_subject,
                                    template,
                                    context),
            'pem_body_text':get_value(cursor,
                                      user,
                                      record_id,
                                      template.def_body_text,
                                      template,
                                      context),
            'pem_body_html':get_value(cursor,
                                      user,
                                      record_id,
                                      template.def_body_html,
                                      template,
                                      context),
            'pem_account_id' :from_account['id'],
            #This is a mandatory field when automatic emails are sent
            'state':'na',
            'folder':'drafts',
            'mail_type':'multipart/alternative'
        }
        #Use signatures if allowed
        if template.use_sign:
            sign = self.pool.get('res.users').read(cursor,
                                                   user,
                                                   user,
                                                   ['signature'],
                                                   context)['signature']
            if sign:
                if mailbox_values['pem_body_text']:
                    mailbox_values['pem_body_text'] += "\n--\n"+sign
                if mailbox_values['pem_body_html']:
                    mailbox_values['pem_body_html'] += sign
        mailbox_id = self.pool.get('poweremail.mailbox').create(
                                                             cursor,
                                                             user,
                                                             mailbox_values,
                                                             context)
        return mailbox_id

    def generate_mail(self,
                      cursor,
                      user,
                      template_id,
                      record_ids,
                      context=None):
        if context is None:
            context = {}
        template = self.browse(cursor, user, template_id, context=context)
        if not template:
            raise Exception("The requested template could not be loaded")
        
        if template.use_filter and template.filter:
            filtered_record_ids=[]
            for record in self.pool.get(template.object_name.model).browse(cursor, user, record_ids, context=context):
                if safe_eval(template.filter, {'o':record, 'self':self, 'cr':cursor, 'context':context}):
                    filtered_record_ids.append(record.id)
            record_ids=filtered_record_ids
        
        report_record_ids = record_ids[:]
        if template.single_email and len(record_ids) > 1:
            # We send a single email for several records
            record_ids = record_ids[:1]

        for record_id in record_ids:
            mailbox_id = self._generate_mailbox_item_from_template(
                                                                cursor,
                                                                user,
                                                                template,
                                                                record_id,
                                                                context)
            mail = self.pool.get('poweremail.mailbox').browse(
                                                        cursor,
                                                        user,
                                                        mailbox_id,
                                                        context=context
                                                              )
            if template.report_template:
                if template.single_email and len(report_record_ids) > 1:
                    # The optional attachment will be generated as a single file for all these records
                    self._generate_attach_reports(
                                              cursor,
                                              user,
                                              template,
                                              report_record_ids,
                                              mail,
                                              context
                                              )
                else:                              
                    self._generate_attach_reports(
                                              cursor,
                                              user,
                                              template,
                                              [record_id],
                                              mail,
                                              context
                                              )
            # Create a partner event
            cursor.execute("SELECT state from ir_module_module where state='installed' and name = 'mail_gateway'")
            mail_gateway = cursor.fetchall()
            if template.partner_event and mail_gateway:
                self._generate_partner_events(cursor,
                                              user,
                                              template,
                                              record_id,
                                              mail,
                                              context)
            # This should be the last statement in this method.
            # This prevents attempts by the scheduler to send
            # Emails before all the work is complete in
            # Generating email, attachments and event
            if not template.save_to_drafts:
                self.pool.get('poweremail.mailbox').write(
                                                    cursor,
                                                    user,
                                                    mailbox_id,
                                                    {'folder':'outbox'},
                                                    context=context
                                                          )
        return True

poweremail_templates()

class poweremail_preview(osv.osv_memory):
    _name = "poweremail.preview"
    _description = "Power Email Template Preview"

    def _get_model_recs(self, cr, uid, context=None):
        if context is None:
            context = {}
        #Fills up the selection box which allows records from the selected object to be displayed
        self.context = context
        if 'active_id' in context.keys() and context.get('active_model', False) == 'poweremail.templates':
            ref_obj_id = self.pool.get('poweremail.templates').read(cr, uid, context['active_id'], ['object_name'], context)['object_name']
            ref_obj_name = self.pool.get('ir.model').read(cr, uid, ref_obj_id[0], ['model'], context)['model']
            ref_obj_ids = self.pool.get(ref_obj_name).search(cr, uid, [], context=context)
            ref_obj_recs = self.pool.get(ref_obj_name).name_get(cr, uid, ref_obj_ids, context)
            return ref_obj_recs

    def _default_model(self, cursor, user, context=None):
        """
        Returns the default value for model field
        @param cursor: Database Cursor
        @param user: ID of current user
        @param context: Open ERP Context
        """
        return self.pool.get('poweremail.templates').read(
                                                   cursor,
                                                   user,
                                                   context['active_id'],
                                                   ['object_name'],
                                                   context)['object_name']

    _columns = {
        'ref_template':fields.many2one(
                                       'poweremail.templates',
                                       'Template', readonly=True),
        'rel_model':fields.many2one('ir.model', 'Model', readonly=True),
        'rel_model_ref':fields.selection(_get_model_recs, 'Referred Document'),
        'to':fields.char('To', size=250, readonly=True),
        'cc':fields.char('CC', size=250, readonly=True),
        'bcc':fields.char('BCC', size=250, readonly=True),
        'subject':fields.char('Subject', size=200, readonly=True),
        'body_text':fields.text('Body', readonly=True),
        'body_html':fields.text('Body', readonly=True),
        'report':fields.char('Report Name', size=100, readonly=True),
    }
    _defaults = {
        'ref_template': lambda self, cr, uid, ctx:ctx['active_id'],
        'rel_model': _default_model
    }

    def on_change_ref(self, cr, uid, ids, rel_model_ref, context=None):
        if context is None:
            context = {}
        if not rel_model_ref:
            return {}
        vals = {}
        if context == {}:
            context = self.context
        template = self.pool.get('poweremail.templates').browse(cr, uid, context['active_id'], context)
        #Search translated template
        lang = get_value(cr, uid, rel_model_ref, template.lang, template, context)
        if lang:
            ctx = context.copy()
            ctx.update({'lang':lang})
            template = self.pool.get('poweremail.templates').browse(cr, uid, context['active_id'], ctx)
        vals['to'] = get_value(cr, uid, rel_model_ref, template.def_to, template, context)
        vals['cc'] = get_value(cr, uid, rel_model_ref, template.def_cc, template, context)
        vals['bcc'] = get_value(cr, uid, rel_model_ref, template.def_bcc, template, context)
        vals['subject'] = get_value(cr, uid, rel_model_ref, template.def_subject, template, context)
        vals['body_text'] = get_value(cr, uid, rel_model_ref, template.def_body_text, template, context)
        vals['body_html'] = get_value(cr, uid, rel_model_ref, template.def_body_html, template, context)
        vals['report'] = get_value(cr, uid, rel_model_ref, template.file_name, template, context)
        return {'value':vals}

poweremail_preview()

class res_groups(osv.osv):
    _inherit = "res.groups"
    _description = "User Groups"
    _columns = {}
res_groups()

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
