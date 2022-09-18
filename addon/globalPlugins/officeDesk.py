# NVDA Office Desk
# Copyright 2022 Joseph Lee and contributors, released under GPL.

# Adds handlers for various controls found across Microsoft Office applications.

import globalPluginHandler
import globalVars
from NVDAObjects.UIA import UIA, SearchField, SuggestionsList, SuggestionListItem
import addonHandler
addonHandler.initTranslation()


# Security: disable the global plugin altogether in secure mode.
def disableInSecureMode(cls):
	return globalPluginHandler.GlobalPlugin if globalVars.appArgs.secure else cls


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if isinstance(obj, UIA):
			# Recognize search field found in backstage view.
			if obj.UIAAutomationId == "HomePageSearchBox":
				clsList.insert(0, SearchField)
			# Also recognize suggestions list and its items.
			elif obj.UIAElement.cachedClassName == "NetUIListView" and isinstance(obj.parent.previous, SearchField):
				clsList.insert(0, SuggestionsList)
			elif obj.UIAElement.cachedClassName == "NetUIListViewItem" and isinstance(obj.parent, SuggestionsList):
				clsList.insert(0, SuggestionListItem)

	def event_UIA_notification(self, obj, nextHandler, displayString=None, activityId=None, **kwargs):
		# In recent versions of Word 365, notification event is used to announce editing functions,
		# some of them being quite anoying.
		if obj.appModule.appName == "winword" and activityId == "AccSN1":
			return
		nextHandler()