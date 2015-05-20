/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.CalendarStories;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class CreateRecurringEventStory extends BaseUserStory {

    private static final String STORY_DESCRIPTION = "Create a recurring event";

    @Override
    public String execute() {
        boolean isStoryComplete;
        String resultMessage;

        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        CalendarSnippets calendarSnippets = new CalendarSnippets(
                getO365MailClient());
        List<String> attendeeEmailAdresses = new ArrayList<>();
        attendeeEmailAdresses.add(GlobalValues.USER_EMAIL);

        try {
            //Create a recurring event
            String newEventId = calendarSnippets.createRecurringCalendarEvent(
                    getStringResource(R.string.calendar_subject_text)
                    , getStringResource(R.string.calendar_body_text)
                    , attendeeEmailAdresses);

            //Cleanup by deleting the event
            calendarSnippets.deleteCalendarEvent(newEventId);
            isStoryComplete = true;
            resultMessage = STORY_DESCRIPTION + ": Recurring event created";
        } catch (ExecutionException | InterruptedException e) {
            isStoryComplete = false;
            resultMessage = FormatException(e, STORY_DESCRIPTION);
        }

        return StoryResultFormatter.wrapResult(resultMessage, isStoryComplete);
    }

    @Override
    public String getDescription() {
        return STORY_DESCRIPTION;
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
