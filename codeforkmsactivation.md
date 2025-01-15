# Microsoft-Office-Product-Activation
Steps toward the MS Office Activation
CMD as Administrator
--------------------------------------------------------------------------------------

cd C:\Program Files\Microsoft Office\Office16
cscript ospp.vbs /sethst:kms.03k.org
cscript ospp.vbs /act

--------------------------------------------------------------------------------------
Step 1. Copy path of MS Office File (C:\Program Files\Microsoft Office\Office16) 
	type "cd"<space>"C:\Program Files\Microsoft Office\Office16" + Enter
Step 2. type "cscript ospp.vbs /sethst:kms.03k.org" + Enter
Step 3. type "cscript ospp.vbs /act" + Enter
	Waiting.... till <<Product activation successful>>
 
**PREVIEWING IN CMD RUN**

Microsoft Windows [Version 10.0.26100.2894]
(c) Microsoft Corporation. All rights reserved.

C:\Windows\System32>cd C:\Program Files\Microsoft Office\Office16

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /sethst:kms.03k.org
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Successfully applied setting.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /act
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Installed product key detected - attempting to activate the following product:
SKU ID: 8d368fc1-9470-4be2-8d66-90e836cbb051
LICENSE NAME: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition
LICENSE DESCRIPTION: Office 24, VOLUME_KMSCLIENT channel
Last 5 characters of installed product key: GCVGB
<Product activation successful>
---------------------------------------
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>
