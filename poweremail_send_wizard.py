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

    def _get_model_recs(self,cr,uid,ctx={}):
        self.context = ctx
        if 'active_id' in ctx.keys():
            ref_obj_id = self.pool.get('poweremail.templates').read(cr,uid,ctx['active_id'],['object_name'])['object_name']
            ref_obj_name = self.pool.get('ir.model').read(cr,uid,ref_obj_id[0],['model'])['model']
            ref_obj_ids = self.pool.get(ref_obj_name).search(cr,uid,[])
            ref_obj_recs = self.pool.get(ref_obj_name).name_get(cr,uid,ref_obj_ids)
            #print ref_obj_recs
            return ref_obj_recs
        
    def _get_accounts(self,cr,uid,ids,ctx={}):
        print cr,uid,ids,ctx
        
    _columns = {
        'ref_template':fields.many2one('poweremail.templates','Template',readonly=True),
        'rel_model':fields.many2one('ir.model','Model',readonly=True),
        'rel_model_ref':fields.selection(_get_model_recs,'Referred Document',readonly=True),
        'from':fields.many2one('poweremail.core_accounts','From Account',),
        'to':fields.char('To',size=100,readonly=True),
        'cc':fields.char('CC',size=100,readonly=True),
        'bcc':fields.char('BCC',size=100,readonly=True),
        'subject':fields.char('Subject',size=200,readonly=True),
        'body_text':fields.text('Body',readonly=True),
        'body_html':fields.text('Body',readonly=True),
        'report':fields.char('Report Name',size=100,readonly=True),
                }
poweremail_send_wizard()
    
    