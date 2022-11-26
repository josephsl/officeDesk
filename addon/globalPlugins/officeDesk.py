# NVDA Office Desk
# Copyright 2022 Joseph Lee and contributors, released under GPL.

# Adds handlers for various controls found across Microsoft Office applications.

import globalPluginHandler
import globalVars
from NVDAObjects.UIA import UIA, SearchField, SuggestionsList, SuggestionListItem
from NVDAObjects.behaviors import EditableTextWithSuggestions
import addonHandler
addonHandler.initTranslation()


# Security: disable the global plugin altogether in secure mode.
def disableInSecureMode(cls):
	return globalPluginHandler.GlobalPlugin if globalVars.appArgs.secure else cls


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		# Most resolved in NVDA 2023.1.
		import versionInfo
		if isinstance(obj, UIA):
			if versionInfo.version_year < 2023:
				# Recognize search field and suggestions list items found in backstage view.
				if obj.UIAAutomationId == "HomePageSearchBox":
					clsList.insert(0, SearchField)
				elif obj.UIAElement.cachedClassName == "NetUIListViewItem" and isinstance(obj.parent, SuggestionsList):
					clsList.insert(0, SuggestionListItem)
			# Also recognize suggestions list.
			if (
				obj.UIAElement.cachedClassName == "NetUIListView"
				and isinstance(obj.parent.previous, (SearchField, EditableTextWithSuggestions))
			):
				clsList.insert(0, SuggestionsList)
