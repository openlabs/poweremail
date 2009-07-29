#########################################################################
#Power Email is a module for Open ERP which enables it to send mails    #
#to customers, suppliers etc. and also has a fiull fledged email client.#
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

{
    "name" : "Powerful Email capabilities for Open ERP",
    "version" : "1.0.2",
    "author" : "Sharoon Thomas, TL-Pragmatic",
    "website" : "",
    "category" : "Added functionality",
    "depends" : ['base'],
    "description": """
    A module similar to the smtpclient and email_sale etc etc. But lot more powerful. Creates three user groups:1.Email Manager(obvious), 2.Email External(Send email to partners),3.Email Internal (mail to seniors etc). the module supports cc, bcc etc which the present smtp client does not. Most unique thing is you can create default settings for sale order, invoice, etc with default cc's,bcc's and default subject, report name and body. the subject, reportname and body takes placeholders which has over 12 functions eg. can get customer name with %(cust_name) etc etc.
        
    """,
    "update_xml": [
        #'security/poweremail_security.xml',
        #'security/ir.model.access.csv',
        'poweremail_workflow.xml',
        'poweremail_core_view.xml',
        'poweremail_template_view.xml',
        'poweremail_send_wizard.xml'
    ],
    "demo_xml" : [],
    "installable": True,
    "active": False,
}

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
