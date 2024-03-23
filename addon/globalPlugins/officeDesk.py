# NVDA Office Desk
# Copyright 2022-2023 Joseph Lee and contributors, released under GPL.

# Adds handlers for various controls found across Microsoft Office applications.

import globalPluginHandler
import globalVars
from NVDAObjects.UIA import UIA, SearchField, SuggestionsList
from NVDAObjects.behaviors import EditableTextWithSuggestions
import addonHandler
addonHandler.initTranslation()


# Security: disable the global plugin altogether in secure mode.
def disableInSecureMode(cls):
	return globalPluginHandler.GlobalPlugin if globalVars.appArgs.secure else cls


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if isinstance(obj, UIA):
			# Recognize suggestions list.
			if (
				obj.UIAElement.cachedClassName == "NetUIListView"
				and isinstance(obj.parent.previous, (SearchField, EditableTextWithSuggestions))
			):
				clsList.insert(0, SuggestionsList)
