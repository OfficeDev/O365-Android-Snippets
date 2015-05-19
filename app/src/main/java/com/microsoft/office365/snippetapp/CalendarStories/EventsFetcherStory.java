/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

package com.microsoft.office365.snippetapp.CalendarStories;

import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Event;

import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Locale;

public class EventsFetcherStory extends BaseUserStory {

    public static final String STORY_DESCRIPTION = "Get calendar events";

    @Override
    public String execute() {
        String returnResult = "";
        if (getO365MailClient() == null) {
            returnResult = "Null OutlookClient";
        }
        try {
            AuthenticationController
                    .getInstance()
                    .setResourceId(
                            getO365MailResourceId());

            CalendarSnippets calendarSnippets = new CalendarSnippets(
                    getO365MailClient());

            //get the calendar events
            List<Event> events = calendarSnippets.getO365Events();

            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append("The following events were retrieved:\n");
            for (Event event : events) {
                sb.append("\t\t" + event.getSubject() + ". " + formatEventDates(event));
                sb.append("\n");
            }
            returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
        } catch (Exception ex) {
            return BaseExceptionFormatter(ex, STORY_DESCRIPTION);
        }
        return returnResult;
    }

    private String formatEventDates(Event thisEvent) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yy - hh:ss a", Locale.US);
        return simpleDateFormat.format(thisEvent.getStart().getTime()).toString();
    }

    @Override
    public String getDescription() {
        return "Gets Events";
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
