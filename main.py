import pymsteams

myTeamsMessage = pymsteams.connectorcard("<Microsoft Webhook URL>")
myTeamsMessage.text("this is my text")
myTeamsMessage.send()