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

class poweremail_templates(osv.osv):
    _name="poweremail.templates"
    _description = 'Power Email Templates for Models'

    _columns = {
        'name' : fields.char('Name of Template',size=100,required=True),
        'object':fields.many2one('ir.model','Model',required=True),
        'def_to':fields.char('Recepient (To)',size=64,required=True),
        'def_cc':fields.char('Default CC',size=64),
        'def_bcc':fields.char('Default BCC',size=64),
        'def_subject':fields.char('Standard Subject',size=200),
        'def_body':fields.text('Standard Body',help="The Signatures will be automatically appended"),
        'use_sign':fields.boolean('Use Signature'),
        'file_name':fields.char('File Name Pattern',size=200,required=True),
        'allowed_groups':fields.one2many('res.groups','poweremail_template',string="Allowed User Groups",  help="Only users from these groups will be allowed to send mails from this ID"),
        'enforce_from_account':fields.many2one('poweremail.core_accounts',string="Enforce From Account",help="Emails will be sent only from this account.",domain="[('company','=','yes')]"),

        'auto_email':fields.boolean('Auto Email'),
        'attached_wkf':fields.many2one('workflow','Workflow'),
        'attached_activity':fields.many2one('workflow.activity','Activity'),
        'server_action':fields.many2one('ir.actions.server','Related Server Action')
        
    }

    _defaults = {

    }
poweremail_templates()
