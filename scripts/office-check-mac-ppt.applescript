-- PowerPoint's JXA save-format enum can silently save PPTX under a .pdf name.
-- Use its native AppleScript terminology, then independently validate PDF bytes.
on run argv
 set sourcePath to item 1 of argv
 set outputPath to item 2 of argv
 set fileNameText to item 3 of argv
 tell application "Microsoft PowerPoint"
  with timeout of 120 seconds
   if exists presentation fileNameText then error "本次核对副本已打开，请关闭它后重试。"
   open (POSIX file sourcePath as alias)
   set deck to presentation fileNameText
   try
    set pageCount to count slides of deck
    set appVersion to version
    save deck in (POSIX file outputPath) as save as PDF
    close deck saving no
    return "{\"openedSuccessfully\":true,\"pdfExported\":true,\"officeVersion\":\"" & appVersion & "\",\"slideCount\":" & pageCount & "}"
   on error errText number errNumber
    try
     close deck saving no
    end try
    error errText number errNumber
   end try
  end timeout
 end tell
end run
