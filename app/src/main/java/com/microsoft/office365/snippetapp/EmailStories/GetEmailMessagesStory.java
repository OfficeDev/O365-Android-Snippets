/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.EmailStories;

import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.services.outlook.Message;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetEmailMessagesStory extends BaseEmailUserStory {

    private static final String STORY_DESCRIPTION = "Gets 10 newest email messages";

    @Override
    public String execute() {
        String returnResult;

        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        try {

            EmailSnippets emailSnippets = new EmailSnippets(
                    getO365MailClient());

            //O365 API called in this helper
            List<Message> messages = emailSnippets.getMailMessages();

            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append("Gets email: " + messages.size() + " messages returned");
            sb.append("\n");
            for (Message m : messages) {
                sb.append("\t\t");
                sb.append(m.getSubject());
                sb.append("\n");
            }
            returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
        } catch (ExecutionException | InterruptedException e) {
            return FormatException(e, STORY_DESCRIPTION);
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
