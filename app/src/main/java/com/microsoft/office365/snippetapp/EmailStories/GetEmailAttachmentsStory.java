/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.EmailStories;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Attachment;
import com.microsoft.outlookservices.FileAttachment;
import com.microsoft.outlookservices.Message;

import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutionException;


public class GetEmailAttachmentsStory extends BaseEmailUserStory {
    private static final String SENT_NOTICE = "Attachment email sent with subject line:";
    private static final boolean IS_INLINE = false;
    private static final String STORY_DESCRIPTION = "Gets the attachments from an email message";

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
            //1. Send an email and store the ID
            String uniqueGUID = java.util.UUID.randomUUID().toString();
            String mailSubject = getStringResource(R.string.mail_subject_text) + uniqueGUID;

            //Add a new email to the user's draft folder
            String emailID = emailSnippets.addDraftMail(GlobalValues.USER_EMAIL,
                    mailSubject,
                    getStringResource(R.string.mail_body_text));

            //Add a text file attachment to the mail added to the draft folder
            emailSnippets.addTextFileAttachmentToMessage(emailID
                    , getStringResource(R.string.text_attachment_contents)
                    , getStringResource(R.string.text_attachment_filename)
                    , IS_INLINE);

            String draftMessageID = emailSnippets.getMailMessageById(emailID).getId();

            //UTC time Immediately before message is sent
            Date sendDate = new Date();
            //Send the draft email to the recipient
            emailSnippets.sendMail(draftMessageID);

            //Get the new message
            Message sentMessage = GetAMessageFromEmailFolder(emailSnippets,
                    getStringResource(R.string.mail_subject_text)
                            + uniqueGUID, getStringResource(R.string.Email_Folder_Inbox));


            StringBuilder sb = new StringBuilder();
            sb.append(SENT_NOTICE);
            sb.append(getStringResource(R.string.mail_subject_text) + uniqueGUID);
            if (sentMessage.getId().length() > 0) {
                List<Attachment> attachments = emailSnippets.getAttachmentsFromEmailMessage(
                        sentMessage.getId());
                //Send the mail with attachments
                //build string for test results on UI
                for (Attachment attachment : attachments) {
                    if (attachment instanceof FileAttachment) {
                        FileAttachment fileAttachment = (FileAttachment) attachment;
                        String fileContents = new String(fileAttachment.getContentBytes(), "UTF-8");
                        sb.append(fileContents);
                        sb.append("/n");
                    }
                }
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
            } else
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), false);


            //3. Delete the email using the ID
            // Boolean result = emailSnippets.deleteMail(emailID);

        } catch (ExecutionException | InterruptedException | UnsupportedEncodingException ex) {
            return FormatException(ex, STORY_DESCRIPTION);
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
