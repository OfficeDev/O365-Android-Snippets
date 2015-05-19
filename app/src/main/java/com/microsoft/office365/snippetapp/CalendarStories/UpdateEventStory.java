/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.CalendarStories;

import android.util.Log;

import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Event;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class UpdateEventStory extends BaseUserStory {


    @Override
    public String execute() {
        //PREPARE
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        CalendarSnippets calendarSnippets = new CalendarSnippets(getO365MailClient());
        List<String> attendeeEmailAddresses = new ArrayList<>();
        attendeeEmailAddresses.add(GlobalValues.USER_EMAIL);
        String newEventId = "";
        //ACT
        try {
            newEventId = calendarSnippets.createCalendarEvent(
                    getStringResource(R.string.calendar_subject_text)
                    , getStringResource(R.string.calendar_body_text)
                    , java.util.Calendar.getInstance()
                    , java.util.Calendar.getInstance()
                    , attendeeEmailAddresses
            );
            Thread.sleep(20000);
            Event updatedEvent = calendarSnippets.updateCalendarEvent(
                    newEventId
                    , getStringResource(R.string.calendar_subject_text)
                            + " Updated Subject"
                    , false
                    , null
                    , null
                    , null
                    , null
            );
            Thread.sleep(20000);
            String updatedSubject = updatedEvent.getSubject();
            //CLEAN UP
            calendarSnippets.deleteCalendarEvent(newEventId);
            //ASSERT
            if (updatedSubject.equals(getStringResource(R.string.calendar_subject_text)
                    + " Updated Subject")) {
                return StoryResultFormatter.wrapResult(
                        "UpdateEventStory: Event "
                                + " updated.", true);
            } else {
                return StoryResultFormatter.wrapResult(
                        "Update Event Story: Update "
                                + " event.", false);
            }
        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update event exception: "
                            + formattedException
                    , false
            );
        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update event story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update event exception: "
                            + formattedException
                    , false
            );
        }
    }

    @Override
    public String getDescription() {
        return "Update a calendar event";
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
