/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.BodyType;
import com.microsoft.outlookservices.Event;
import com.microsoft.outlookservices.Importance;
import com.microsoft.outlookservices.ItemBody;

import java.util.Calendar;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetFilteredImportantEvents extends BaseUserStory {

    @Override
    public String execute() {
        boolean isSucceeding;
        AuthenticationController
                .getInstance()
                .setResourceId(getO365MailResourceId());
        CalendarSnippets calendarSnippets = new CalendarSnippets(getO365MailClient());

        try {
            //Set up one important event to test with
            Event testEvent = new Event();
            testEvent.setSubject(getStringResource(R.string.calendar_subject_text));

            //Set body on test event
            ItemBody itemBody = new ItemBody();
            itemBody.setContent(getStringResource(R.string.calendar_body_text));
            itemBody.setContentType(BodyType.HTML);
            testEvent.setBody(itemBody);

            //Set start and end time for event
            Calendar eventTime = Calendar.getInstance();
            testEvent.setStart(eventTime);
            eventTime.add(Calendar.HOUR_OF_DAY, 2);
            testEvent.setIsAllDay(false);
            testEvent.setEnd(eventTime);
            testEvent.setImportance(Importance.High);

            //Create test event on tenant
            testEvent = calendarSnippets.createCalendarEvent(testEvent);

            //Retrieve important events (should include our test event)
            List<Event> importantEvents = calendarSnippets.getImportantEvents();

            //Check that all events are important to determine if story succeeded.
            isSucceeding = true;
            for (Event event : importantEvents) {
                if (event.getImportance() != Importance.High) {
                    isSucceeding = false;
                    break;
                }
            }

            //Delete test event from tenant
            calendarSnippets.deleteCalendarEvent(testEvent.getId());

        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("EventFilter", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Filter important events: " + formattedException
                    , false
            );
        }
        if (isSucceeding) {
            return StoryResultFormatter.wrapResult("FilterImportantEventsStory: Important events found.", true);
        } else {
            return StoryResultFormatter.wrapResult("FilterImportantEventsStory: Important events not found.", false);
        }
    }

    @Override
    public String getDescription() {
        return "Gets events filtered by most important";
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
