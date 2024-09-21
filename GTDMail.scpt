-- ====================================================================
-- GTDMail: GTD tickler file functionality for Mail.app

(* ====================================================================   
 * BSD 3-Clause License
 * 
 * Copyright (c) 2005, Zak Greant <zak@greant.com>
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 * 
 * 1. Redistributions of source code must retain the above copyright notice, this
 *    list of conditions and the following disclaimer.
 * 
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 *    this list of conditions and the following disclaimer in the documentation
 *    and/or other materials provided with the distribution.
 * 
 * 3. Neither the name of the copyright holder nor the names of its
 *    contributors may be used to endorse or promote products derived from
 *    this software without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
 * FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
 * DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 * CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
 * OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
 * OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *)

-- ====================================================================
-- script properties: names for the various tickler file folders

property FOLDER_INBOX : "@INBOX" -- where messages from the tickler file are dumped
property FOLDER_TICKLER : "@TICKLER" -- the main tickler mailbox
property FOLDER_MONTH : "Month"
property FOLDERS_MONTHS : {"01.Jan", "02.Feb", "03.Mar", "04.Apr", "05.May", "06.Jun", "07.Jul", "08.Aug", "09.Sep", "10.Oct", "11.Nov", "12.Dec"}
property FOLDERS_WEEKS : {"Week 1", "Week 2", "Week 3", "Week 4", "Week 5"}
property FOLDERS_DAYS : {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31}

-- ====================================================================
-- run handlers

on run
	my tickleMain()
end run

using terms from application "Mail"
	on perform mail action with messages
		my tickleMain()
	end perform mail action with messages
end using terms from

-- ====================================================================
-- Main routine

on tickleMain()
	tell application "Mail"
		-- Check to see if the @Tickler and @INBOX mail folders are present; if not create them.
		try
			set mbox to mailbox named FOLDER_INBOX
			set mbox to mailbox named FOLDER_TICKLER
			my tickleEmails()
		on error
			my setupMailboxes()
		end try
	end tell
end tickleMain

-- ====================================================================
-- subroutines

-- Move tickled emails to FOLDER_INBOX
on tickleEmails()
	tell application "Mail"
		-- set the mail folders for the current day, week and month
		set currentDay to item (day of (current date)) of FOLDERS_DAYS
		set currentWeek to item (((day of (current date)) - 1) div 7 + 1) of FOLDERS_WEEKS
		set CurrentMonth to item (month of (current date) as integer) of FOLDERS_MONTHS
		
		-- move message from the current day, week and month folders to folder FOLDER_INBOX
		my moveMessages(FOLDER_TICKLER & "/" & FOLDER_MONTH & "/" & CurrentMonth)
		my moveMessages(FOLDER_TICKLER & "/" & currentWeek)
		my moveMessages(FOLDER_TICKLER & "/" & currentWeek & "/" & currentDay)
		
		-- flag and mark all tickler messages as unread
		my flagAndUnreadAllMessages({mailbox named FOLDER_TICKLER, mailbox named FOLDER_INBOX})
	end tell
end tickleEmails

-- move messages from the tickler mailboxes to FOLDER_INBOX
on flagAndUnreadAllMessages(mboxes)
	tell application "Mail"
		repeat with mbox in mboxes
			set read status of every message of mbox to false
			set flagged status of every message of mbox to true
			
			repeat with submbox in mailboxes of mbox
				my flagAndUnreadAllMessages(submbox)
			end repeat
		end repeat
	end tell
end flagAndUnreadAllMessages

-- move messages from the tickler mailboxes to FOLDER_INBOX
on moveMessages(mbox)
	tell application "Mail"
		try
			move (every message of mailbox (mbox)) to (mailbox named FOLDER_INBOX)
		end try
	end tell
end moveMessages

-- Create the tickler mailboxes
on setupMailboxes()
	display dialog "Can't find '" & FOLDER_TICKLER & "' folder and its associated subfolders; would you like me to create them?"
	tell application "Mail"
		make new mailbox with properties {name:FOLDER_INBOX}
		-- Don't create a mailbox for FOLDER_TICKLER or FOLDER_MONTH. This prevents us from accidentally losing any messages in these folders.
		repeat with folder in FOLDERS_MONTHS
			make new mailbox with properties {name:(FOLDER_TICKLER & "/" & FOLDER_MONTH & "/" & folder)}
		end repeat
		repeat with thisWeek in FOLDERS_WEEKS
			make new mailbox with properties {name:(FOLDER_TICKLER & "/" & thisWeek)}
		end repeat
		repeat with thisDay in FOLDERS_DAYS
			set thisWeek to item ((thisDay - 1) div 7 + 1) of FOLDERS_WEEKS
			make new mailbox with properties {name:(FOLDER_TICKLER & "/" & thisWeek & "/" & thisDay)}
		end repeat
	end tell
	display dialog "The '" & FOLDER_TICKLER & "' folder and subfolders have been created."
end setupMailboxes
