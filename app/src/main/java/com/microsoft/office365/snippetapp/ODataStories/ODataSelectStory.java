/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.ODataStories;

import android.util.Log;

import com.microsoft.office365.snippetapp.Snippets.ODataSystemQuerySnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.services.outlook.Message;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class ODataSelectStory extends BaseUserStory {
    private static final String STORY_DESCRIPTION = "Use $select to reduce payload when getting messages";

    @Override
    public String execute() {
        boolean isStoryComplete = false;
        StringBuilder returnResult = new StringBuilder();

        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());
        try {
            ODataSystemQuerySnippets oDataSystemQuerySnippets = new ODataSystemQuerySnippets();

            //O365 API called in this helper
            List<Message> messages = oDataSystemQuerySnippets.getMailMessagesUsing$select(getO365MailClient());

            returnResult.append(STORY_DESCRIPTION)
                    .append(": ")
                    .append(messages.size())
                    .append(" messages returned")
                    .append("\n");

            for (Message message : messages) {
                returnResult.append("\t\tFrom: ")
                        .append(message.getFrom())
                        .append("\n\t\tIs Read: ")
                        .append(message.getIsRead().toString())
                        .append("\n\t\tSubject: ")
                        .append(message.getSubject())
                        .append("\n");
            }
            isStoryComplete = true;
        } catch (ExecutionException | InterruptedException e) {
            isStoryComplete = false;
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Get email story", formattedException);
            returnResult = new StringBuilder();
            returnResult.append(STORY_DESCRIPTION)
                    .append(": ")
                    .append(formattedException);
        }
        return StoryResultFormatter.wrapResult(returnResult.toString(), isStoryComplete);
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
