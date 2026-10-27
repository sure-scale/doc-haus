-- A short-lived launcher. Terminal keeps ownership of the product stack.
use framework "AppKit"
use framework "ScriptingBridge"
use scripting additions

global lastStartRequestedAt, unavailableBrowsers

on run
    set lastStartRequestedAt to missing value
    set unavailableBrowsers to {}
    my openDocHaus()
    my finishOrWait()
end run

on reopen
    my openDocHaus()
    my finishOrWait()
end reopen

on open location launchURL
    -- A URL event can be the first event delivered to a newly started applet.
    try
        set initializedAt to lastStartRequestedAt
        set initializedBrowsers to unavailableBrowsers
    on error
        set lastStartRequestedAt to missing value
        set unavailableBrowsers to {}
    end try
    if my isLaunchURL(launchURL) then
        my openDocHaus()
    else
        try
            -- The fixed native helper owns the strict Word URI parser. The
            -- supplied URI is only a single shell-quoted argument to it.
            set launcher to POSIX path of (path to resource "open-word.sh")
            do shell script "/bin/bash " & quoted form of launcher & " " & quoted form of launchURL
        on error
            display dialog "doc.haus could not open this Word link. Return to the document in doc.haus and choose Open in Word again." with title "doc.haus" buttons {"OK"} default button "OK" with icon caution
        end try
    end if
    my finishOrWait()
end open location

on isLaunchURL(launchURL)
    if class of launchURL is not text then return false
    considering case
        return launchURL is "doc.haus://launch" or launchURL is "doc.haus://launch/"
    end considering
end isLaunchURL

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
            if family is not "" and browserID is not in unavailableBrowsers then
                set instances to current application's NSRunningApplication's runningApplicationsWithBundleIdentifier:browserID
                set browserInstance to my selectBrowserInstance(instances as list, appURL)
                if browserInstance is not missing value then
                    try
                        if my focusExistingTab(browserInstance, family, targetOrigin) then return
                    on error
                        -- Denied or failed events fall back without retrying this
                        -- browser during the launcher's short lifetime.
                        set end of unavailableBrowsers to browserID
                    end try
                end if
            end if
        end if
    end try
    -- Unknown/ambiguous browsers, missing tabs and denied access open normally.
    do shell script "/usr/bin/open " & quoted form of targetOrigin
end openBrowser

-- A directly spawned browser can share a bundle ID with the user's desktop
-- browser. LaunchServices supplies launchDate only for its own app launches.
-- Do not query any browser when more than one desktop instance is eligible.
on selectBrowserInstance(instances, expectedAppURL)
    set chosen to missing value
    set expectedURL to expectedAppURL's URLByResolvingSymlinksInPath()'s URLByStandardizingPath()
    repeat with instanceReference in instances
        set instance to contents of instanceReference
        if not (instance's isTerminated() as boolean) and (instance's activationPolicy() as integer) is 0 and instance's launchDate() is not missing value then
            set candidateURL to instance's bundleURL()
            if candidateURL is not missing value then
                set candidateURL to candidateURL's URLByResolvingSymlinksInPath()'s URLByStandardizingPath()
                if (candidateURL's isEqual:expectedURL) as boolean then
                    if chosen is not missing value then return missing value
                    set chosen to instance
                end if
            end if
        end if
    end repeat
    return chosen
end selectBrowserInstance

on browserInstanceIsLive(instance)
    if instance's isTerminated() as boolean then return false
    set currentInstance to current application's NSRunningApplication's runningApplicationWithProcessIdentifier:(instance's processIdentifier())
    if currentInstance is missing value then return false
    return (instance's isEqual:currentInstance) as boolean
end browserInstanceIsLive

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

on focusExistingTab(browserInstance, family, targetOrigin)
    if not my browserInstanceIsLive(browserInstance) then return false
    -- Every browser event is addressed to this process, never its bundle ID.
    set browser to current application's SBApplication's applicationWithProcessIdentifier:(browserInstance's processIdentifier())
    if browser is missing value then return false
    browser's setTimeout:300 -- Apple event timeout is measured in 60 Hz ticks.
    if not my focusBrowserTab(browser, family, targetOrigin) then return false
    if not my browserInstanceIsLive(browserInstance) then return false
    return (browserInstance's activateWithOptions:(current application's NSApplicationActivateIgnoringOtherApps)) as boolean
end focusExistingTab

on focusBrowserTab(browser, family, targetOrigin)
    set usableWindows to {}
    set urlsByWindow to {}
    set currentByWindow to {}
    set browserWindows to browser's |windows|()'s |get|() as list
    repeat with windowReference in browserWindows
        set browserWindow to contents of windowReference
        -- Chromium's private windows are excluded before reading any tab URLs.
        if family is not "chromium" or (browserWindow's |mode|() as text) is "normal" then
            set tabObjects to browserWindow's |tabs|()'s |get|() as list
            set tabAddresses to {}
            repeat with tabReference in tabObjects
                set tabObject to contents of tabReference
                set end of tabAddresses to tabObject's |URL|() as text
            end repeat
            set end of usableWindows to browserWindow
            copy tabAddresses to end of urlsByWindow
            if family is "chromium" then
                set end of currentByWindow to browserWindow's activeTabIndex() as integer
            else
                set end of currentByWindow to browserWindow's currentTab()'s |index|() as integer
            end if
        end if
    end repeat
    set chosen to my matchingTab(urlsByWindow, currentByWindow, targetOrigin)
    if item 1 of chosen is 0 then return false
    set targetWindow to item (item 1 of chosen) of usableWindows
    set tabIndex to item 2 of chosen
    if family is "chromium" then
        targetWindow's setActiveTabIndex:tabIndex
        targetWindow's setMinimized:false
    else
        set targetTabs to targetWindow's |tabs|()'s |get|() as list
        targetWindow's setCurrentTab:(item tabIndex of targetTabs)
        targetWindow's setMiniaturized:false
    end if
    targetWindow's setIndex:1
    return true
end focusBrowserTab

on quit
    -- Closing this Dock entry never stops Terminal or its services.
    continue quit
end quit
