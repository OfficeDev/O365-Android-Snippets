/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Message;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class RespondToCalendarEventInviteStory extends BaseUserStory {

    public static final int MAX_TRY_COUNT = 20;

    @Override
    public String execute() {
        //PREPARE
        String returnValue = StoryResultFormatter.wrapResult("Create Event story", false);

        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        CalendarSnippets calendarSnippets = new CalendarSnippets(
                getO365MailClient());
        EmailSnippets emailSnippets = new EmailSnippets(
                getO365MailClient());

        List<String> attendeeEmailAddresses = new ArrayList<>();
        attendeeEmailAddresses.add(GlobalValues.USER_EMAIL);
        String newEventId = "";

        try {
        String uniqueGUID = java.util.UUID.randomUUID().toString();
        String subjectLine = getStringResource(R.string.calendar_subject_text)
                + ":"
                + uniqueGUID;
         try {

             //Store the date and time that the email is sent in UTC
             Date sentDate = new Date();
            newEventId = calendarSnippets.createCalendarEvent(
                    subjectLine
                    , getStringResource(R.string.calendar_body_text)
                    , java.util.Calendar.getInstance()
                    , java.util.Calendar.getInstance()
                    , attendeeEmailAddresses);

            Thread.sleep(20000);
            if (calendarSnippets.respondToCalendarEventInvite(newEventId
                    , GlobalValues.USER_EMAIL, getStringResource(R.string.CalendarEvent_Accept)) != null){



                Thread.sleep(20000);
                String attendeeStatus = calendarSnippets.getEventAttendeeStatus(
                        newEventId
                        , GlobalValues.USER_EMAIL);

                //Get the new message
                String emailId = "";
                int tryCount = 0;

                calendarSnippets.deleteCalendarEvent(newEventId);
                //Try to get the newly sent email event invitation
                // from user's Sent Items folder at least once.
                //continue trying to get the email while the email is not found
                //and the loop has tried less than MAX_TRY_COUNT times.
                do {
                    List<Message> mailIds = emailSnippets
                            .GetMailboxMessagesByFolderName_Subject(
                                    subjectLine
                                    , getStringResource(R.string.Email_Folder_Sent));
                    if (mailIds.size() > 0) {
                        for (Message message : mailIds) {
                            emailId = message.getId();
                            emailSnippets.deleteMail(emailId);
                        }
                    }
                    tryCount++;

                    //Stay in loop while these conditions are true.
                    //If either condition becomes false, break
                } while (emailId.length() == 0 && tryCount < MAX_TRY_COUNT);

                if (attendeeStatus.toLowerCase().contains("accept")) {
                    return StoryResultFormatter.wrapResult(
                            "Respond to event invite story: Event "
                                    + " response.", true);
                } else {
                    return StoryResultFormatter.wrapResult(
                            "Respond to event invite story: Event "
                                    + " response failed. " + attendeeStatus, false
                    );

                }
            } else {
                return StoryResultFormatter.wrapResult(
                        "Respond to event invite story: Event "
                                + " is null.", false
                );

            }
        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Respond to event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Respond to event exception: "
                            + formattedException
                    , false
            );

        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Respond to event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Respond to event exception: "
                            + formattedException
                    , false
            );
        }

    }

    @Override
    public String getDescription() {
        return "Responds to an event invite";
    }

}
// *********************************************************
//
// O365-Android-Snippets, https://github.com/OfficeDev/O365-Android-Snippets
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
