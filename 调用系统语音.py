import win32com.client as win

Speak = win.Dispatch("SAPI.SpVoice")
Speak.Speak('哈哈哈哈哈哈')