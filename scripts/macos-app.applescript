-- A small stay-open launcher. Terminal keeps ownership of the product stack.
global lastStartRequestedAt

on run
    set lastStartRequestedAt to missing value
    my openDocHaus()
end run

on reopen
    my openDocHaus()
end reopen

on openDocHaus()
    try
        set action to "launch"
        if lastStartRequestedAt is not missing value then
            if (current date) - lastStartRequestedAt < 10 then set action to "waiting"
        end if
        set launcher to POSIX path of (path to resource "open-doc-haus.sh")
        set outcome to do shell script "/bin/bash " & quoted form of launcher & " " & quoted form of action
        if outcome is "started" then set lastStartRequestedAt to current date
    on error messageText
        display dialog messageText with title "doc.haus" buttons {"OK"} default button "OK" with icon caution
    end try
end openDocHaus

on quit
    -- Closing this Dock entry never stops Terminal or its services.
    continue quit
end quit
