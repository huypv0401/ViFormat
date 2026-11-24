-- Hop thoai Unicode cho ViFormat (VBA tren Mac goi qua AppleScriptTask).
on ShowMsg(txt)
	tell application "System Events" to activate
	display dialog txt with title "ViFormat" buttons {"OK"} default button 1 with icon note
	return "ok"
end ShowMsg

-- Hien van ban dai co thanh cuon (dong bat dau bang # la tieu de, in dam)
on ShowText(txt)
	set tmp to (system attribute "TMPDIR") & "viformat_msg.txt"
	set fh to open for access POSIX file tmp with write permission
	set eof of fh to 0
	write txt to fh as «class utf8»
	close access fh
	set here to do shell script "dirname " & quoted form of (POSIX path of (path to me))
	do shell script "osascript -l JavaScript " & quoted form of (here & "/ViFormat.js") & " " & quoted form of tmp
	return "ok"
end ShowText
