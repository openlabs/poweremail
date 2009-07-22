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
import pygtk
pygtk.require('2.0')
import gtk

from osv import osv, fields

class poweremail_templates(osv.osv):
    _name="poweremail.templates"
    _description = 'Power Email Templates for Models'

    _columns = {
        'name' : fields.char('Name of Template',size=100,),
        'object':fields.many2one('ir.model','Model'),
        'def_to':fields.char('Recepient (To)',size=64,),
        'def_cc':fields.char('Default CC',size=64),
        'def_bcc':fields.char('Default BCC',size=64),
        'def_subject':fields.char('Standard Subject',size=200),
        'def_body':fields.text('Standard Body',help="The Signatures will be automatically appended"),
        'use_sign':fields.boolean('Use Signature'),
        'file_name':fields.char('File Name Pattern',size=200,),
        'allowed_groups':fields.one2many('res.groups','poweremail_template',string="Allowed User Groups",  help="Only users from these groups will be allowed to send mails from this ID"),
        'enforce_from_account':fields.many2one('poweremail.core_accounts',string="Enforce From Account",help="Emails will be sent only from this account.",domain="[('company','=','yes')]"),

        'auto_email':fields.boolean('Auto Email'),
        'attached_wkf':fields.many2one('workflow','Workflow'),
        'attached_activity':fields.many2one('workflow.activity','Activity'),
        'server_action':fields.many2one('ir.actions.server','Related Server Action'),
        'model_object_field':fields.many2one('ir.model.fields',string="Available Options"),
        'sub_object':fields.many2one('ir.model','Sub Object'),
        'sub_model_object_field':fields.many2one('ir.model.fields','Sub Field'),
        'null_value':fields.char('Null Value',help="This Value is used if the field is empty",size=50),
        'copyvalue':fields.char('Placeholder',size=100)
    }

    _defaults = {

    }
    
    def _field_changed(self,cr,uid,ids,parent_field):
        print "Parent:",parent_field
        if parent_field:
            field_obj = self.pool.get('ir.model.fields').browse(cr,uid,parent_field)
            print field_obj.ttype
            if field_obj.ttype in ['many2one','one2many','many2many']:
                res_ids=self.pool.get('ir.model').search(cr,uid,[('model','=',field_obj.relation)])
                print res_ids[0]
                if res_ids:
                    print self.write(cr,uid,ids,{'sub_object':res_ids[0]})
                    return {'value':{'sub_object':res_ids[0]}}
                else:
                    return {'value':{'sub_object':False,'sub_model_object_field':False}}
            else:
                    return {'value':{'sub_object':False,'sub_model_object_field':False}}
        else:
                    return {'value':{'sub_object':False,'sub_model_object_field':False}}
        
    def add_field(self,cr,uid,ids,ctx={}):
        clipboard = gtk.clipboard_get()
        if self.read(cr,uid,ids,['model_object_field'])[0]['model_object_field']:
            obj_id = self.read(cr,uid,ids,['model_object_field'])[0]['model_object_field'][0]
            obj_br = self.pool.get('ir.model.fields').browse(cr,uid,obj_id)
            obj_not = obj_br.name
            if self.read(cr,uid,ids,['sub_model_object_field'])[0]['sub_model_object_field']:
                obj_id = self.read(cr,uid,ids,['sub_model_object_field'])[0]['sub_model_object_field'][0]
                obj_br = self.pool.get('ir.model.fields').browse(cr,uid,obj_id)
                obj_not = obj_not + "." + obj_br.name
                if self.read(cr,uid,ids,['null_value'])[0]['null_value']:
                    obj_not = obj_not + "/" + self.read(cr,uid,ids,['null_value'])[0]['null_value']
            obj_not = "[[$." + obj_not + "]]"
            return {'value':{'copyvalue':obj_not}}
            clipboard.set_text(obj_not)
            clipboard.store()
            print "Clipboard:",clipboard.wait_for_text()

        
poweremail_templates()

class res_groups(osv.osv):
    _inherit = "res.groups"
    _columns={
        'poweremail_template':fields.many2one('poweremail.templates'),
    }
res_groups()