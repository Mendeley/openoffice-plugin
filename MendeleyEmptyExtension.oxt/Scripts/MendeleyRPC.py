# ***** BEGIN LICENSE BLOCK *****
#
# Copyright (c) 2009 Mendeley Ltd.
# Copyright (c) 2006 Center for History and New Media
#                    George Mason University, Fairfax, Virginia, USA
#                    http://chnm.gmu.edu
#
# Licensed under the Educational Community License, Version 1.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.opensource.org/licenses/ecl1.php
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# ***** END LICENSE BLOCK *****
#
# Thanks to the Zotero developers whose Word/Open Office plugin source code was
# frequently referred to and borrowed from in the development of this plugin
#
# author: steve.ridout@mendeley.com

import unohelper
import httplib

from com.sun.star.task import XJob

g_ImplementationHelper = unohelper.ImplementationHelper()

class MendeleyRPC(unohelper.Base, XJob):
    #The component must have a ctor with the component context as argument
    def __init__(self,ctx):
        self.closed = 0
        self.ctx = ctx
        
    def execute(self, args):
        q = args[0].Value
        q=q.encode('utf-8')
        headers = {"Content-Type": "utf8/xml"}
        h1=httplib.HTTPConnection("localhost:5002")
        h1.request("POST","",q,headers)
        response=h1.getresponse()
        data=response.read()
        data=data.decode('utf-8')
        #h1.close()
        return data

g_ImplementationHelper.addImplementation(MendeleyRPC, "org.openoffice.pyuno.MendeleyRPC", ("com.sun.star.task.MendeleyRPC",),)
