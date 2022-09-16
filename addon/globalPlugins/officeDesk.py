# NVDA Office Desk
# Copyright 2022 Joseph Lee and contributors, released under GPL.

# Adds handlers for various controls found across Microsoft Office applications.

import globalPluginHandler
import globalVars
import addonHandler
addonHandler.initTranslation()


# Security: disable the global plugin altogether in secure mode.
def disableInSecureMode(cls):
	return globalPluginHandler.GlobalPlugin if globalVars.appArgs.secure else cls


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def event_UIA_notification(self, obj, nextHandler, displayString=None, activityId=None, **kwargs):
		# In recent versions of Word 365, notification event is used to announce editing functions,
		# some of them being quite anoying.
		if obj.appModule.appName == "winword" and activityId == "AccSN1":
			return
		nextHandler()
