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
import com.microsoft.outlookservices.ResponseType;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class RespondToCalendarEventInviteStory extends BaseEmailUserStory {

    @Override
    public String execute() {
        //PREPARE
        boolean isStoryComplete;
        String resultMessage = "";
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
        String uniqueGUID = java.util.UUID.randomUUID().toString();
        String subjectLine = getStringResource(R.string.calendar_subject_text)
                + ":"
                + uniqueGUID;
        try {
            String newEventId = calendarSnippets.createCalendarEvent(
                    subjectLine
                    , getStringResource(R.string.calendar_body_text)
                    , java.util.Calendar.getInstance()
                    , java.util.Calendar.getInstance()
                    , attendeeEmailAddresses);

            Thread.sleep(5000);//wait for server to send event invitation
            if (calendarSnippets.respondToCalendarEventInvite(newEventId
                    , GlobalValues.USER_EMAIL, ResponseType.Accepted) != null) {
                Thread.sleep(5000);//wait for server to update attendee status in event
                ResponseType attendeeStatus = calendarSnippets.getEventAttendeeStatus(
                        newEventId
                        , GlobalValues.USER_EMAIL);

                //Validate the attendee status was set to accepted as expected
                if (attendeeStatus == ResponseType.Accepted) {
                    isStoryComplete = true;
                    resultMessage = "Respond to event invite story: Event response.";
                } else {
                    isStoryComplete = false;
                    resultMessage = "Respond to event invite story: Event response failed. "
                            + attendeeStatus;
                }

                //CLEANUP by cancelling event
//                calendarSnippets.respondToCalendarEventInvite(newEventId
//                        , GlobalValues.USER_EMAIL, ResponseType.Declined);
                calendarSnippets.deleteCalendarEvent(newEventId);
//                Message message = GetAMessageFromEmailFolder(emailSnippets, subjectLine, getStringResource(R.string.Email_Folder_Sent));
//                if (message!=null){
//                    emailSnippets.deleteMail(message.getId());
//                }

            } else {
                isStoryComplete = false;
                resultMessage = "Respond to event invite story: Event is null.";
            }
        } catch (ExecutionException | InterruptedException e) {
            isStoryComplete = false;
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            resultMessage = "Respond to event exception: "
                    + formattedException;
            e.printStackTrace();
            Log.e("Respond to event story", formattedException);
        }
        return StoryResultFormatter.wrapResult(resultMessage, isStoryComplete);
    }

    @Override
    public String getDescription() {
        return "Responds to accept an event invite";
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
