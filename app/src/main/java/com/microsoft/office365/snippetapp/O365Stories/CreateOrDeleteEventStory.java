/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.CalendarSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryAction;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

//This story handles both of the following stories that appear in the UI list
// based on strings passed in the constructor...
//- Create an event (which is then deleted for cleanup)
//- Delete an event (which is created first and then deleted)
public class CreateOrDeleteEventStory extends BaseUserStory {
    private final String CREATE_DESCRIPTION = "Adds a new calendar event";
    private final String CREATE_TAG = "Create event story";
    private final String CREATE_SUCCESS = "CreateEventStory: Event created.";
    private final String CREATE_ERROR = "Create event exception: ";
    private final String DELETE_DESCRIPTION = "Deletes a calendar event";
    private final String DELETE_TAG = "Delete event story";
    private final String DELETE_SUCCESS = "DeleteEventStory: Event deleted.";
    private final String DELETE_ERROR = "Delete event exception: ";

    private String mDescription;
    private String mLogTag;
    private String mSuccessDescription;
    private String mErrorDescription;

    public CreateOrDeleteEventStory(StoryAction action) {
        switch (action) {
            case CREATE: {
                mDescription = CREATE_DESCRIPTION;
                mLogTag = CREATE_TAG;
                mSuccessDescription = CREATE_SUCCESS;
                mErrorDescription = CREATE_ERROR;
                break;
            }
            case DELETE: {
                mDescription = DELETE_DESCRIPTION;
                mLogTag = DELETE_TAG;
                mSuccessDescription = DELETE_SUCCESS;
                mErrorDescription = DELETE_ERROR;
                break;
            }
        }
    }

    @Override
    public String execute() {
        //PREPARE
        AuthenticationController
                .getInstance()
                .setResourceId(
                        super.getO365MailResourceId());

        CalendarSnippets calendarSnippets = new CalendarSnippets(
                getO365MailClient());

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
                    , attendeeEmailAddresses);

            //Delete event
            calendarSnippets.deleteCalendarEvent(newEventId);
        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e(mLogTag, formattedException);
            return StoryResultFormatter.wrapResult(
                    mErrorDescription + formattedException
                    , false
            );
        }
        return StoryResultFormatter.wrapResult(mSuccessDescription, true);
    }

    @Override
    public String getDescription() {
        return mDescription;
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
