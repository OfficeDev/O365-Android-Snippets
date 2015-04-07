/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class AcceptEventInviteStory extends BaseUserStory {
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

        List<String> attendeeEmailAddresses = new ArrayList<>();
        attendeeEmailAddresses.add(GlobalValues.USER_EMAIL);
        String newEventId = "";
        //ACT
        try {
            newEventId = calendarSnippets.createCalendarEvent(
                    "Subject"
                    , "<p class=MsoNormal>Hello world!</p>"
                    , java.util.Calendar.getInstance()
                    , java.util.Calendar.getInstance()
                    , attendeeEmailAddresses);

            String addedEventId = calendarSnippets.getCalendarEventId(newEventId);
            calendarSnippets.acceptCalendarEventInvite(
                    newEventId
                    , GlobalValues.USER_EMAIL);


            String attendeeStatus = calendarSnippets.getEventAttendeeStatus(
                    newEventId
                    , GlobalValues.USER_EMAIL);
            //CLEAN UP
            calendarSnippets.deleteCalendarEvent(newEventId);
            //ASSERT
            if (attendeeStatus.toLowerCase().contains("accept")) {
                return StoryResultFormatter.wrapResult(
                        "Accept event invite story: Event "
                                + " accepted.", true
                );

            } else {
                return StoryResultFormatter.wrapResult(
                        "Accept event invite story: Event "
                                + " accepted.", false
                );

            }
        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Accept event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Accept event exception: "
                            + formattedException
                    , false
            );

        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Accept event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Accept event exception: "
                            + formattedException
                    , false
            );
        }

    }

    @Override
    public String getDescription() {
        return "Accept event invite";
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
