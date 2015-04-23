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
import com.microsoft.outlookservices.Event;

import java.util.Date;
import java.util.List;

public class SendEmailWithCalendarAttachStory extends BaseEmailUserStory {
    public static final String STORY_DESCRIPTION = "Sends an email message with a calendar event attachment";
    public static final String SENT_NOTICE = "Email sent with subject line:";
    EmailSnippets mEmailSnippets;

    @Override
    public String execute() {
        String returnResult = "";
        try {

            AuthenticationController
                    .getInstance()
                    .setResourceId(
                            getO365MailResourceId());

            EmailSnippets emailSnippets = new EmailSnippets(
                    getO365MailClient());
            mEmailSnippets = emailSnippets;

            CalendarSnippets calendarSnippets = new CalendarSnippets(getO365MailClient());
            List<Event> eventsToAttach = calendarSnippets.getO365Events();

            //Store the date and time that the email is sent in UTC
            Date sentDate = new Date();
            //1. Send an email and store the ID
            String uniqueGUID = java.util.UUID.randomUUID().toString();

            //Create a new email message but do not send yet
            String newEmailId = emailSnippets.addDraftMail(
                    GlobalValues.USER_EMAIL
                    , getStringResource(R.string.mail_subject_text) + uniqueGUID
                    , getStringResource(R.string.mail_body_text));

            if (eventsToAttach.size() > 0) {

                //Attach email message to new draft email
                emailSnippets.addItemAttachment(newEmailId, eventsToAttach.get(1));

                //Send draft email
                emailSnippets.sendMail(newEmailId);

                DeleteAMessageFromMailFolder(emailSnippets,
                        getStringResource(R.string.mail_subject_text)
                                + uniqueGUID,
                        getStringResource(R.string.Email_Folder_Draft));
                DeleteAMessageFromMailFolder(emailSnippets,
                        getStringResource(R.string.mail_subject_text)
                                + uniqueGUID,
                        getStringResource(R.string.Email_Folder_Sent));

                returnResult = StoryResultFormatter.wrapResult(
                        STORY_DESCRIPTION, true
                );
            }
        }
        catch (Exception ex) {
            String formattedException = APIErrorMessageHelper.getErrorMessage(ex.getMessage());
            Log.e("Send msg w/ event ", formattedException);
            returnResult = StoryResultFormatter.wrapResult(
                    "Send mail exception: "
                            + formattedException
                    , false
            );
        }

        return returnResult;

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