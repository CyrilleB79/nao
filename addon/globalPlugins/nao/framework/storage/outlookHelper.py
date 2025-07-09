#Nao (NVDA Advanced OCR) is an addon that improves the standard OCR capabilities that NVDA provides on modern Windows versions.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Last update 2025-07-09
#Copyright (C) 2024-2025 Alessandro Albano, Davide De Carne, Simone Dal Maso, Cyrille Bougot

import os
import tempfile
import api
from logHandler import log
from comtypes import COMError

CONTROL_ID_ATTACHMENTS = 4306

# MAPI property: PR_ATTACHMENT_HIDDEN
# https://learn.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor.getproperty
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

class OutlookHelper:
	def is_outlook(obj=None):
		if obj is None: obj = api.getForegroundObject()
		return obj and obj.appModule and obj.appModule.appName and obj.appModule.appName == 'outlook'

	def __init__(self, obj=None):
		self.obj = obj
		self.nativeOm = None
		if OutlookHelper.is_outlook(obj):
			self.nativeOm = obj.appModule.nativeOm

	def is_valid(self):
		return self.nativeOm != None

	def is_active(self):
		return self.handle and self.handle == ctypes.windll.user32.GetForegroundWindow()

	def focusInAttachmentsList(self):
		for o in reversed(api.getFocusAncestors()):
			if getattr(o, 'windowControlID') == CONTROL_ID_ATTACHMENTS:
				return True
		return False		

	def indexOfAttachment(self, obj):
		for idx, o in enumerate(obj.parent.children):
			if obj == o:
				return idx
		raise LookupError('Unable to find index of current item')

	@staticmethod
	def isInlineAttachment(attachment):
		try:
			return attachment.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN)
		except COMError:
			return False

	def currentFileWithPath(self):
		obj = api.getFocusObject()
		if not self.focusInAttachmentsList():
			log.debug('Not in attachments list')
			return None, None
		index = self.indexOfAttachment(obj)
		inspector = self.nativeOm.ActiveInspector()
		tempDir = ''
		if inspector is not None:
			tempDir = tempfile.mkdtemp()
			mailItem = inspector.CurrentItem
			visibleAttachments = []
			for i in range(1, mailItem.Attachments.Count + 1):
				att = mailItem.Attachments.Item(i)
				if not self.isInlineAttachment(att):
					visibleAttachments.append(att)
			attachment = visibleAttachments[index]
			tempPath = os.path.join(tempDir, attachment.FileName)
			attachment.SaveAsFile(tempPath)
			log.debug(f"{tempPath=}")
			return tempPath, tempDir
		return None, None
