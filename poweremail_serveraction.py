#########################################################################
#Power Email is a module for Open ERP which enables it to send mails    #
#The server action which allows sending from workflows is coded here    #
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
#__author__="sharoon"
#__date__ ="$4 Aug, 2009 1:47:16 PM$"
from osv import fields,osv
import netsvc

class actions_server(osv.osv):
    _inherit = 'ir.actions.server'
    _description = 'Server action with Power Email update'
    _columns = {
        'state': fields.selection([
            ('poweremail','Power Email'),
            ('client_action','Client Action'),
            ('dummy','Dummy'),
            ('loop','Iteration'),
            ('code','Python Code'),
            ('trigger','Trigger'),
            ('email','Email'),
            ('sms','SMS'),
            ('object_create','Create Object'),
            ('object_write','Write Object'),
            ('other','Multi Actions'),
        ], 'Action Type', required=True, size=32, help="Type of the Action that is to be executed"),
        'poweremail_template':fields.many2one('poweremail.templates','Template',)#In view customize such that domain('object_name','=',model_id)
    }

    def run(self, cr, uid, ids, context={}):
        print cr,uid,ids,context
        logger = netsvc.Logger()
        logger.notifyChannel('Server Action', netsvc.LOG_INFO, 'Started Server Action with Power Email update')

        for action in self.browse(cr, uid, ids, context):
            if action.state=='poweremail':
                if not action.poweremail_template:
                    raise osv.except_osv(_('Error'), _("Please specify an template to use for auto email in poweremail !"))
                templ_id = action.poweremail_template.id
                
                self.pool.get('poweremail.templates').generate_mail(cr,uid,templ_id,[context['active_id']])
                return False
            else:
                return super(actions_server,self).run(cr, uid, ids, context)
actions_server()