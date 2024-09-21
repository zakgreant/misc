-- ====================================================================
-- email2ical: create a iCal event from one or more Mail.app email messages. See below for important details.

(* ====================================================================   
 * Copyright 2006 Foo Associates Inc. (written by J. A. (Zak) Greant <zak@greant.com>)
 * All rights reserved. This code is licensed under the Modified/New BSD License.
 * 
 *  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the
 * following conditions are met:
 *  * Redistributions of source code must retain the above copyright notice, this list of conditions and the following 
 *     disclaimer.
 *  * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the 
 *     following disclaimer in the documentation and/or other materials provided with the distribution.
 *  * Neither the name of Foo Associates Inc. nor the names of its contributors may be used to endorse or promote 
 *     products derived from this software without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS 
 * OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF 
 * MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE 
 * COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE
 * GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED 
 * AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
 * OF THE POSSIBILITY OF SUCH DAMAGE.
 *)

(* ====================================================================
 * email2iCal Overview
 * Creates one iCal event for every selected email message in Mail.app
 * 
 * The summary of the event will be taken from the subject of the email message, while the description of the 
 * event will be taken from the body of the message.
 *
 * Each event is created EventOffset after the current date and time (rounded down to the current hour) and 
 * lasts for EventDuration
 *
 * Event alarms are created at AlarmOne and AlarmTwo minutes before or after the event (use negative integers 
 * for an alarm before the event start)
 *
 * Customize EmailAddress2calendar. If entries in EmailAddress2calendar can be matched to one of the email 
 * addresses the message was delivered to, then the event will be created in the corresponding calendar.
 * Note: Partial matches are valid.
 *
 * Set DefaultCalendar to a reasonable value that can be used if an address is not found in
 * EmailAddress2calendar
 *)

-- TODO
-- Put the unique ID of the email message in the calender event and write script to lookup the
-- corresponding email in Mail.app
-- 
-- Cook up a personal version that archives the email or trashes it

-- ====================================================================
-- Script properties
property EmailAddress2calendar : {
	{address:"@ez.no", calendarName:"eZ systems"},
	{address:"@lists.ez.no", calendarName:"eZ systems"},
	{address:"zak@fooassociates.com", calendarName:"Foo Associates"},
	{address:"zak@mozillafoundation.org", calendarName:"MoFo"}
		}
property DefaultCalendar : "Zak"
property EventOffset : 24 * hours
property EventDuration : hours
property AlarmOne : -60 -- the time in minutes for the first alarm (a leading minus means "before the event")
property AlarmTwo : -1440 -- the time in minutes for the second alarm

-- ====================================================================
-- Handlers for running the script or invoking it via the script menu or via a mail rule
on run
	my email2ical()
end run

-- Handler when script is run from either the script menu in Mail or triggered by a rule
using terms from application "Mail"
	on perform mail action with messages
		my email2ical()
	end perform mail action with messages
end using terms from


-- ====================================================================
-- Main routine

on email2ical()
	tell application "Mail"
		repeat with currentViewer in message viewers
			if (count of selected messages of currentViewer) is greater than 0 then
				try
					my makeIcalEvents(selected messages of currentViewer)
				on error
					display dialog "Something went wrong"
				end try
				
				tell application "iCal" to activate
			else
				display dialog "No messages selected in current Message Viewer"
			end if
		end repeat
	end tell
end email2ical

-- ====================================================================
-- subroutines

-- make all iCal events
on makeIcalEvents(allMessages)
	tell application "Mail"
		repeat with currentMessage in allMessages
			set eventSummary to subject of currentMessage
			set eventDescription to content of currentMessage
			set deliveryAddress to address of recipient of currentMessage
			set currentCalendar to DefaultCalendar -- this will be overwritten, if we can find a match in th repeat loop below
			
			tell application "iCal"
				set exitBothRepeats to false
				repeat with pair in EmailAddress2calendar
					repeat with currentAddress in deliveryAddress
						if currentAddress contains address of pair then
							set currentCalendar to calendarName of pair
							set exitBothRepeats to true
							exit repeat
						end if
					end repeat
					if exitBothRepeats is true then
						exit repeat
					end if
				end repeat
				
				set cal to (first item of (calendars whose name is currentCalendar))
				set newEvent to make new event at end of cal
				
				tell newEvent
					set summary to eventSummary
					set description to eventDescription
					
					set baseDate to (current date)
					set minutes of baseDate to 0
					set seconds of baseDate to 0
					
					set start date to baseDate + EventOffset
					set end date to baseDate + EventOffset + EventDuration
					try
						make new display alarm at end of display alarms with properties {trigger interval:AlarmOne}
						make new display alarm at end of display alarms with properties {trigger interval:AlarmTwo}
					end try
					show
				end tell
			end tell
		end repeat
	end tell
end makeIcalEvents
