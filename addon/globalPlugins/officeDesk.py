# NVDA Office Desk
# Copyright 2022 Joseph Lee and contributors, released under GPL.

# Adds handlers for various controls found across Microsoft Office applications.

import globalPluginHandler
import addonHandler
addonHandler.initTranslation()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def event_UIA_notification(self, obj, nextHandler, displayString=None, activityId=None, **kwargs):
		# In recent versions of Word 365, notification event is used to announce editing functions,
		# some of them being quite anoying.
		if obj.appModule.appName == "winword" and activityId == "AccSN1":
			return
		nextHandler()
