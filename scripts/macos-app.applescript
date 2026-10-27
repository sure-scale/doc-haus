-- A short-lived launcher. Terminal keeps ownership of the product stack.
use framework "AppKit"
use scripting additions

global lastStartRequestedAt, deniedBrowsers

on run
    set lastStartRequestedAt to missing value
    set deniedBrowsers to {}
    my openDocHaus()
    my finishOrWait()
end run

on reopen
    my openDocHaus()
    my finishOrWait()
end reopen

on idle
    return my finishOrWait()
end idle

-- Only a cold launch needs a brief lifetime to absorb duplicate opens before
-- Terminal has claimed its runtime. A ready browser interface exits promptly.
on launchGraceRemaining(requestedAt, nowAt)
    if requestedAt is missing value then return 0
    set remaining to 10 - (nowAt - requestedAt)
    if remaining < 0 then return 0
    return remaining
end launchGraceRemaining

on finishOrWait()
    set remaining to my launchGraceRemaining(lastStartRequestedAt, current date)
    if remaining is 0 then quit
    return remaining
end finishOrWait

on openDocHaus()
    try
        set action to "launch"
        if lastStartRequestedAt is not missing value then
            if (current date) - lastStartRequestedAt < 10 then set action to "waiting"
        end if
        set launcher to POSIX path of (path to resource "open-doc-haus.sh")
        set outcome to do shell script "/bin/bash " & quoted form of launcher & " " & quoted form of action
        if outcome is "started" then set lastStartRequestedAt to current date
        if outcome starts with "ready " then
            my openBrowser(text 7 thru -1 of outcome)
            set lastStartRequestedAt to missing value
        end if
    on error messageText
        display dialog messageText with title "doc.haus" buttons {"OK"} default button "OK" with icon caution
        set lastStartRequestedAt to missing value
    end try
end openDocHaus

-- No browser is scripted until the dispatcher has verified runtime ownership.
-- These events are sent by this applet, keeping its Automation consent honest.
on openBrowser(targetOrigin)
    try
        set targetURL to current application's NSURL's URLWithString:targetOrigin
        set appURL to current application's NSWorkspace's sharedWorkspace()'s URLForApplicationToOpenURL:targetURL
        if appURL is not missing value then
            set browserID to (current application's NSBundle's bundleWithURL:appURL)'s bundleIdentifier() as text
            set family to my browserFamily(browserID)
            if family is not "" and browserID is not in deniedBrowsers then
                -- Bundle-ID Apple Events cannot choose between separate browser
                -- instances. Leave isolated/other-agent instances untouched.
                set instances to current application's NSRunningApplication's runningApplicationsWithBundleIdentifier:browserID
                if my canReuseBrowserInstanceCount(instances's |count|() as integer) then
                    try
                        with timeout of 5 seconds
                            if my focusExistingTab(browserID, family, targetOrigin) then return
                        end timeout
                    on error number errorNumber
                        -- A denial falls back without repeatedly asking this session.
                        if errorNumber is -1743 then set end of deniedBrowsers to browserID
                    end try
                end if
            end if
        end if
    end try
    -- Unknown/ambiguous browsers, missing tabs and denied access open normally.
    do shell script "/usr/bin/open " & quoted form of targetOrigin
end openBrowser

on canReuseBrowserInstanceCount(instanceCount)
    return instanceCount is 1
end canReuseBrowserInstanceCount

on browserFamily(browserID)
    if browserID is in {"com.apple.Safari", "com.apple.SafariTechnologyPreview"} then return "safari"
    if browserID is in {"com.google.Chrome", "com.google.Chrome.beta", "com.google.Chrome.dev", "com.google.Chrome.canary", "org.chromium.Chromium", "com.microsoft.edgemac", "com.brave.Browser", "com.vivaldi.Vivaldi"} then return "chromium"
    return ""
end browserFamily

on matchesOrigin(candidateURL, targetOrigin)
    if class of candidateURL is not text then return false
    if candidateURL is targetOrigin then return true
    return candidateURL starts with (targetOrigin & "/") or candidateURL starts with (targetOrigin & "?") or candidateURL starts with (targetOrigin & "#")
end matchesOrigin

-- Prefer an already-selected matching tab before changing any tab selection.
-- Pure selection over URL metadata; it never returns or persists that metadata.
on matchingTab(urlsByWindow, currentByWindow, targetOrigin)
    repeat with windowIndex from 1 to count urlsByWindow
        set tabURLs to item windowIndex of urlsByWindow
        set currentIndex to item windowIndex of currentByWindow
        if currentIndex > 0 and currentIndex <= (count tabURLs) then
            if my matchesOrigin(item currentIndex of tabURLs, targetOrigin) then return {windowIndex, currentIndex}
        end if
    end repeat
    repeat with windowIndex from 1 to count urlsByWindow
        set tabURLs to item windowIndex of urlsByWindow
        repeat with tabIndex from 1 to count tabURLs
            if my matchesOrigin(item tabIndex of tabURLs, targetOrigin) then return {windowIndex, tabIndex}
        end repeat
    end repeat
    return {0, 0}
end matchingTab

on focusExistingTab(browserID, family, targetOrigin)
    set usableWindows to {}
    set urlsByWindow to {}
    set currentByWindow to {}
    tell application id browserID
        set browserWindows to (get every window) as list
        repeat with windowReference in browserWindows
            set browserWindow to contents of windowReference
            if family is "chromium" then
                -- Raw dictionary codes avoid requiring Chrome at compile time.
                -- Ignore incognito windows before asking for any tab URLs.
                if (get «class mode» of browserWindow) is "normal" then
                    set tabObjects to (get every «class CrTb» of browserWindow) as list
                    set tabAddresses to {}
                    repeat with tabReference in tabObjects
                        set tabObject to contents of tabReference
                        set end of tabAddresses to (get «class URL » of tabObject) as text
                    end repeat
                    set end of usableWindows to browserWindow
                    copy tabAddresses to end of urlsByWindow
                    set end of currentByWindow to (get «class acTI» of browserWindow)
                end if
            else
                set tabObjects to (get every «class bTab» of browserWindow) as list
                set tabAddresses to {}
                repeat with tabReference in tabObjects
                    set tabObject to contents of tabReference
                    set end of tabAddresses to (get «class pURL» of tabObject) as text
                end repeat
                set end of usableWindows to browserWindow
                copy tabAddresses to end of urlsByWindow
                set end of currentByWindow to (get «class pidx» of «class cTab» of browserWindow)
            end if
        end repeat
    end tell
    set chosen to my matchingTab(urlsByWindow, currentByWindow, targetOrigin)
    if item 1 of chosen is 0 then return false
    set targetWindow to item (item 1 of chosen) of usableWindows
    set tabIndex to item 2 of chosen
    tell application id browserID
        if family is "chromium" then
            set «class acTI» of targetWindow to tabIndex
        else
            set «class cTab» of targetWindow to item tabIndex of (every «class bTab» of targetWindow)
        end if
        set «class pmnd» of targetWindow to false
        set index of targetWindow to 1
        activate
    end tell
    return true
end focusExistingTab

on quit
    -- Closing this Dock entry never stops Terminal or its services.
    continue quit
end quit
