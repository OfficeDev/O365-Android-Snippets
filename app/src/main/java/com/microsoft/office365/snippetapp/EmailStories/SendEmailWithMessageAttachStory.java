/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.EmailStories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Message;

import java.util.Date;

public class SendEmailWithMessageAttachStory extends BaseEmailUserStory {

    public static final String STORY_DESCRIPTION = "Sends an email message with a message attachment";
    public static final String SENT_NOTICE = "Email sent with subject line:";
    public static final boolean IS_INLINE = false;


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

            //Store the date and time that the email is sent in UTC
            Date sentDate = new Date();
            //1. Send an email and store the ID
            String uniqueGUID = java.util.UUID.randomUUID().toString();
            String emailID = emailSnippets.createAndSendMail(
                    GlobalValues.USER_EMAIL
                    , getStringResource(R.string.mail_subject_text)
                            + uniqueGUID, getStringResource(R.string.mail_body_text));


            Message messageToAttach = GetAMessageFromEmailFolder(emailSnippets,
                    getStringResource(R.string.mail_subject_text)
                            + uniqueGUID, getStringResource(R.string.Email_Folder_Inbox));

            if (messageToAttach != null) {
                //Create a new email message but do not send yet
                String newEmailId = emailSnippets.addDraftMail(
                        GlobalValues.USER_EMAIL
                        , getStringResource(R.string.mail_subject_text) + uniqueGUID
                        , getStringResource(R.string.mail_body_text));

                //Attach email message to new draft email
                emailSnippets.addItemAttachment(
                        newEmailId
                        , messageToAttach
                        , IS_INLINE);

                //Send draft email
                emailSnippets.sendMail(newEmailId);

                DeleteAMessageFromMailFolder(emailSnippets,
                        getStringResource(R.string.mail_subject_text)
                                + uniqueGUID,
                        getStringResource(R.string.Email_Folder_Sent));

                returnResult = StoryResultFormatter.wrapResult(
                        STORY_DESCRIPTION, true
                );
            }


        } catch (Exception ex) {
            String formattedException = APIErrorMessageHelper.getErrorMessage(ex.getMessage());
            Log.e("Send msg w/ message ", formattedException);
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
