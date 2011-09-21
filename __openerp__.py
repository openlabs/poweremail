#########################################################################
#                                                                       #
# Copyright (C) 2010-2011 Openlabs Technologies & Consulting (P) LTD    #
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
    "name" : "Poweremail",
    "version" : "0.7",
    "author" : "Sharoon Thomas, Openlabs",
    "website" : "http://openlabs.co.in/article/poweremail",
    "category" : "Added functionality",
    "depends" : [
        'base',
    ],
    "description": """
Power Email - extends the most Power ful open source ERP with email 
which powers the world today.

Features:

1. Multiple Email Accounts
2. Company & Personal Email accounts
3. Security (In Progress)
4. Email Folders (Inbox.Outbox.Drafts etc)
5. Sending of Mails via SMTP (SMTP SSL also supported)
6. Reception of Mails (IMAP & POP3) With SSL & Folders for IMAP supported
7. Simple one point Template designer which automatically updates system. 
    No extra config req.
8. Automatic Email feature on workflow stages

NOTE: This is a beta release. Please update bugs at:

https://github.com/openlabs/poweremail/issues/new

    """,
    "init_xml": [
        'poweremail_scheduler_data.xml'
    ],
    "update_xml": [
        'security/poweremail_security.xml',
        'security/ir.model.access.csv',
        'poweremail_workflow.xml',
        'poweremail_core_view.xml',
        'poweremail_template_view.xml',
        'poweremail_send_wizard.xml',
        'poweremail_mailbox_view.xml',
        'poweremail_serveraction_view.xml',
    ],
    "installable": True,
    "active": False,
}
