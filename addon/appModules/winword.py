# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2019-2020 NV Access Limited, Cyrille Bougot
# Extended by Joseph Lee in 2022 as part of Office Desk add-on

""" App module for Microsoft Word.
Word and Outlook share a lot of code and components. This app module gathers the code that is relevant for
Microsoft Word only.
"""

from nvdaBuiltin.appModules.winword import AppModule, WinwordWordDocument
import api
from NVDAObjects.behaviors import Dialog
from NVDAObjects.IAccessible.winword import WordDocument as IAccessibleWordDocument
from NVDAObjects.UIA.wordDocument import WordDocument as UIAWordDocument


class AppModule(AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if UIAWordDocument in clsList or IAccessibleWordDocument in clsList:
			clsList.insert(0, WinwordWordDocument)

	def event_UIA_notification(self, obj, nextHandler, activityId=None, **kwargs):
		# NVDA Core issue 10950: in recent versions of Word 365, notification event is used to
		# announce editing functions, some of them being quite anoying.
		if activityId == "AccSN1":
			return
		nextHandler()


class WinwordWordDocument(WinwordWordDocument):

	def _get_name(self):
		# Some places in Word have actual labels, not document title.
		# NVDA Core issue 14156: notably, envelopes dialog edit fields do have usable labels.
		name = super(WinwordWordDocument, self).name
		if isinstance(self, UIAWordDocument) and isinstance(api.getForegroundObject(), Dialog):
			name = self.UIAElement.currentName
		return name
