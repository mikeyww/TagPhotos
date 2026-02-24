version := '1.0'

/* TagPhotos ---------------------------------------------------------------------------------------------------
This script modifies photographs' IPTC description tags.
Key features:
• Fast and easy to use
• Quickly append a photo's creation date to the description
• Select multiple photos to tag in sequence
• See a preview of each photo alongside its description tag
By mikeyww in U.S.A. • For AutoHotkey version 2.0.20
25 Feb 2026 (v1.0) : Initial release
Known issue:
 • Clicking a disabled menu item activates the menu such that pressing a menu shortcut letter without ALT then
   activates the corresponding menu item
https://github.com/mikeyww/TagPhotos
https://www.autohotkey.com/boards/
----------------------------------------------------------------------------------------------------------------
*/
#Requires AutoHotkey 2
ITEM           := Map(                                                            ; Menu-bar items
                     'File'   , [1, '&File `tF2'   , select ]
                   , 'Date'   , [2, '&Date `tF3'   , addDate]
                   , 'Update' , [3, '&Update `tF4' , write  ]
                   , 'Refresh', [4, '&Refresh `tF5', read   ]
                   , 'Prev'   , [5, '&Prev `tPgUp' , nav    ]
                   , 'Next'   , [6, '&Next `tPgDn' , nav    ]
                   , 'Help'   , [7, '&Help `tF1' , helpShow ]
                   , 'About'  , [8, '&About'       , about  ]
                  )
OUT            := 'tag.txt'
CRLF           := '`r`n'
MBAR           := []
PICWIDTH       := 400                                                             ; Picture width
MBAR.Length    := ITEM.Count
ICON           := 'picture_icon-icons.com_71126-2.ico'                            ; Icon for this program
APP            := 'exiftool.exe'                                                  ; Manage tags: https://exiftool.org/
SOUND          := A_WinDir '\Media\Windows Information Bar.wav'                   ; Sound to play upon writing metadata
TAG            := 'Caption-Abstract'                                              ; IPTC description tag
SCRIPT         := StrReplace(A_ScriptName, '.ahk')                                ; Used in file selection dialog
GUITITLE       := SCRIPT                                                          ; Prefix for GUI title
CHANGED        := 'FFFF9E'                                                        ; Yellow for GUI background
SAVED          := '88FF88'                                                        ; Green for GUI background
WS_EX_TOPMOST  := 262144                                                          ; For MsgBox
URL            := 'https://github.com/mikeyww/'                                   ; TagPhotos Web site
helpText       := Map()                                                           ; Used to set the help GUI's edit control
aboutStr       := SCRIPT                                              '`n`n'
                . 'Version ' version                                  '`n`n'
                . 'AutoHotkey version: ' A_AhkVersion                 '`n`n'
                . 'Process: '            A_AhkPath                    '`n`n'
                . 'Icons with updated fill:`nhttps://icon-icons.com/' '`n`n'
                . 'ExifTool by Phil Harvey:`nhttps://exiftool.org/'   '`n`n'
                . 'Copyright 2026 mikeyww (from AutoHotkey forum)'    '`n`n'
                . 'https://github.com/mikeyww/'
tagTxt         := (str, tag) => RegExMatch(str, tag '\s*:\s*\K.+', &m) ? m[] : '' ; Extract tag value from metadata
toEnd          := (guiCtrl) => (guiCtrl.Focus(), Sleep(25), Send('^{End}'))       ; Navigate to end of field
g2             := Gui(, 'Please wait')                                            ; Progress GUI for loading photos
Try TraySetIcon ICON
For k, v in ITEM
 MBAR[v[1]] := k
g2.SetFont 's18'
g2.BackColor   := CHANGED
status         := g2.AddText('w400 Center')
prog           := g2.AddProgress('wp cBlue')
g2.OnEvent 'Close', (gui) => Reload()

; Text for the help GUI
helpTextArr := [
   ['Introduction', SCRIPT ' is a simple photograph description tagger for JPG files.'                    '`n`n'
                 . 'The program reads and writes the photos`' "Caption-Abstract" (description) IPTC tag.' '`n`n'
                 . 'Multiple photos can be selected for reading or writing.'                              '`n`n'
                 . 'The photo`'s creation date is displayed and can be appended to the description.'      '`n`n'
                 . 'The PageUp and PageDown keys can be used to navigate to the previous or next photo.'
   ]
 , ['File'   , 'Select photographs to tag.']
 , ['Date'   , 'Append the photo`'s creation date to the description.']
 , ['Update' , 'Write the description to the photo file`'s IPTC metadata.'          '`n`n'
             . 'The metadata will not be saved unless this menu item is activated.' '`n`n'
             . 'After writing the metadata, the program will re-read the metadata.'
   ]
 , ['Refresh', 'Read the description from the photo file`'s IPTC metadata.']
 , ['Prev'   , 'Navigate to the previous photo.' '`n`n'
             . 'Navigation is disabled until the current photo`'s metadata have been read.'
   ]
 , ['Next'   , 'Navigate to the next photo.' '`n`n'
             . 'Navigation is disabled until the current photo`'s metadata have been read.'
   ]
 , ['About'  , aboutStr]
]
For arr in helpTextArr
 helpText[arr[1]] := arr[2]                 ; Help text, by topic

; GUI for help
help := Gui('+AlwaysOnTop +Resize -DPIScale', 'Help for ' SCRIPT)
help.OnEvent 'Size', help_Size              ; When the help GUI is resized
help.SetFont 's10'
help.BackColor := '4F90DB' ; Blue
help.OnEvent 'Escape', (gui) => gui.Hide()
help_LV  := help.AddListView('NoSort -Multi w150 r10 BackgroundF0F0F0', ['Topic'])
help_LV.OnEvent 'ItemFocus', help_ItemFocus ; When a ListView item is newly focused
For arr in helpTextArr
 help_LV.Add , arr[1]                       ; Add the help topic to the ListView
help_ed := help.AddEdit('ym r30 BackgroundWhite ReadOnly')
OnExit (exitReason, exitCode) => (FileExist(OUT) && FileRecycle(OUT))
FileExist(APP) || (MsgBox('File not found. Aborting.`n`n' APP, 'Error', 'Icon!'), ExitApp())

; ==================================
select ; Select photos to load
; ==================================

select(itemName?, itemPos?, m?) {  ; F2 = Select photos to update
 static FILEMUSTEXIST := 1
 static TABLEN        := 25        ; Maximum tab text length
 static WIDTH         := 500       ; Width of right panel
 global selected, g, tab, picPath, desc, created
 If !(sel := FileSelect('M' FILEMUSTEXIST,, SCRIPT ' - Select JPG image files', 'JPG image files (*.jpg)')).Length
  Return
 selected := sel
 Try g.Destroy
 g := Gui(, GUITITLE)
 g.SetFont 's10'
 g.MenuBar := MenuBar()
 For i in MBAR                                                                       ; Populate the menu bar
  If i != 'Prev' && i != 'Next' || selected.Length > 1
   g.MenuBar.Add ITEM[i][2], ITEM[i][3]
 tab := g.AddTab3('-Background -Wrap Buttons')                                       ; Create tab control
 status.Text := ''
 (selected.Length > 1) && g2.Show()
 For k, image in selected {                                                          ; Add one photo per new tab
  status.Text := 'Loading #' k ' of ' selected.Length
  prog.Value  := 100 * k / selected.Length
  SplitPath(image,,,, &fnBare), tab.Add([SubStr(fnBare, 1, TABLEN)])                 ; Use shortened file name as new tab's name
  tab.UseTab(k), g.AddPic('h-1 Section w' PICWIDTH, image)                           ; Add a photo to the new tab
 }
 g2.Hide(), tab.UseTab()                                                             ; Use no tab for subsequent controls
 g.AddText 'x+m ys w' WIDTH, 'File:'
   picPath := g.AddEdit('wp ReadOnly')                                               ; Path to photo file
 g.AddText 'wp', 'Description:'
   desc    := g.AddEdit('wp r10')                                                    ; Description of photo
 g.AddText 'y+12', 'Created:'
   created := g.AddEdit('x+m yp-4 w100 ReadOnly')                                    ; Date photo was created
 tab.OnEvent('Change', (tb, info) => read()), desc.OnEvent('Change', desc_Change)    ; When tab or description changes
 g.OnEvent('Escape', (gui) => ExitApp()), read(), g.Show()
}

addDate(itemName, itemPos, m) {                              ; F3 = Append photo creation date to description
 If created.Text {
  desc.Text .= created.Text
  toEnd(desc), desc_Change(desc)
 }
}

write(itemName, itemPos, m) {                                ; F4 = Write the photo description
 ; ExifTool preserves a copy of the original file by adding "_original" to the copy's file name
 If FileExist(picPath.Text) {
  RunWait APP ' -' StrReplace(TAG, ' ') '="' RegExReplace(desc.Text, '\R+', ' ') '" "' picPath.Text '"',, 'Hide'
  read(), FileExist(SOUND) && SoundPlay(SOUND)               ; Read the photo description & creation date
 }
}

read(itemName?, itemPos?, m?) {                              ; F5 = Read the photo description & creation date
 g.Title := GUITITLE ' - #' tab.Value ' of ' selected.Length ; Update GUI title
 If FileExist(picPath.Text := selected[tab.Value]) {
  For k, v in ITEM                                           ; Disable all menu items
   Try g.MenuBar.Disable v[2]
  g.BackColor   := CHANGED                                   ; Change the background color to yellow
  desc.Text     := created.Text := ''                        ; Clear the fields for description & date created
  RunWait A_ComSpec ' /c ' APP ' "' picPath.Text '">' OUT,, 'Hide' ; Read the photo's metadata
  If FileExist(OUT) {
   txt := FileRead(OUT)                                      ; Get the metadata as text
   desc.Text    := tagTxt(txt, TAG)                          ; Populate the description field
   created.Text := RegExMatch(tagTxt(txt, 'Create Date'), '(\d{4}):(\d\d):(\d\d)', &part) ? part[2] '/' part[3] '/' part[1] : ''
   toEnd desc                                                ; Focus on the description field
   g.BackColor  := SAVED                                     ; Green
   For k, v in ITEM                                          ; Enable all menu items except "Update"
    If k != 'Update'
     Try g.MenuBar.Enable(v[2])
  }
 }
}

nav(itemName, itemPos?, m?) {                                  ; F11 or F12 = Navigate to previous or next photo
 If InStr(itemName, 'Prev')
       tab.Value := tab.Value = 1 ? selected.Length : tab.Value - 1
 Else  tab.Value := tab.Value = selected.Length ? 1 : tab.Value + 1
 read
}

about(itemName, itemPos, m) {
 If 'Yes' = MsgBox(aboutStr '`n`nVisit the Web site?', 'About ' SCRIPT, 'Iconi YNC Default2 ' WS_EX_TOPMOST)
  Run URL
}

desc_Change(e, info?) {              ; Photo description changed
 If e.Gui.BackColor != CHANGED {
  e.Gui.BackColor   := CHANGED       ; Change background color to yellow
  g.MenuBar.Enable ITEM['Update'][2] ; Enable the "Update" menu item
 }
}

helpShow(itemName, itemPos, m) {                           ; Show the help GUI
 help_LV.Modify 1, 'Focus Select'                          ; Select the first topic
 help_ItemFocus help_LV, 1                                 ; Display corresponding text in edit control
 g.GetPos(, &y), g.GetClientPos(,,, &h)                    ; Get main GUI's y-position & client height
 help.Show 'x20 y' y ' w' 0.4 * A_ScreenWidth ' h' h + 30  ; Help GUI's height will match main GUI's height
 help_LV.Focus
}

help_ItemFocus(LV, item) {  ; Set the edit control's contents based on the newly selected item
 topic := LV.GetText(item)
 help_ed.Text := StrUpper(topic) StrReplace('`n`n' helpText[topic], '`n', CRLF)
}

help_Size(gui, minMax, w, h) {                            ; The help GUI was resized
 h -= 2 * gui.MarginY                                     ; New height of controls
 help_LV.GetPos ,, &LVwidth                               ; Get width of ListView
 help_ed.Move   ,, w - 2 * gui.MarginX - LVwidth - 15, h  ; Adjust edit control's dimensions
 help_ed.Redraw                                           ; Redraw edit control to account for GUI's new size
 help_LV.Move  ,,, h                                      ; Adjust ListView's height
}