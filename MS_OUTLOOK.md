# Microsoft Outlook recipes

## Get unread messages from Inbox

```applescript
tell application "Microsoft Outlook"
  return incoming messages of inbox whose is read is false
end tell
```
