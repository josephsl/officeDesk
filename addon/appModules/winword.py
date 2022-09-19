# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2019-2020 NV Access Limited, Cyrille Bougot
# Extended by Joseph Lee in 2022 as part of Office Desk add-on

""" App module for Microsoft Word.
Word and Outlook share a lot of code and components. This app module gathers the code that is relevant for
Microsoft Word only.
"""

import appModuleHandler
from scriptHandler import script
import ui
import api
from NVDAObjects.behaviors import Dialog
from NVDAObjects.IAccessible.winword import WordDocument as IAccessibleWordDocument
from NVDAObjects.UIA.wordDocument import WordDocument as UIAWordDocument
from NVDAObjects.window.winword import WordDocument
from NVDAObjects.UIA import UIA


class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if UIAWordDocument in clsList or IAccessibleWordDocument in clsList:
			clsList.insert(0, WinwordWordDocument)


class WinwordWordDocument(WordDocument):

	def _get_name(self):
		# Some places in Word have actual labels, not document title.
		# NVDA Core issue 14156: notably, envelopes dialog edit fields do have usable labels.
		name = super(WinwordWordDocument, self).name
		if isinstance(self, UIA) and isinstance(api.getForegroundObject(), Dialog):
			name = self.UIAElement.currentName
		return name

	@script(gesture="kb:control+shift+e")
	def script_toggleChangeTracking(self, gesture):
		if not self.WinwordDocumentObject:
			# We cannot fetch the Word object model, so we therefore cannot report the status change.
			# The object model may be unavailable because it's within Windows Defender Application Guard.
			# In this case, just let the gesture through and don't report anything.
			return gesture.send()
		val = self._WaitForValueChangeForAction(
			lambda: gesture.send(),
			lambda: self.WinwordDocumentObject.TrackRevisions
		)
		if val:
			# Translators: a message when toggling change tracking in Microsoft word
			ui.message(_("Change tracking on"))
		else:
			# Translators: a message when toggling change tracking in Microsoft word
			ui.message(_("Change tracking off"))
