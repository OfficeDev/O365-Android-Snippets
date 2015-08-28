/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
package com.microsoft.office365.snippetapp.EmailStories;

import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.outlookservices.Message;
import java.util.List;
import java.util.concurrent.ExecutionException;

public abstract class BaseEmailUserStory extends BaseUserStory {

    private static final int MAX_POLL_REQUESTS = 20;
    public static final int THREAD_SLEEP_TIME = 3000;

    public abstract String execute();

    public abstract String getDescription();

    /**
     * Gets first message found with the given subject line from the user's inbox.
     * Uses a polling technique because typically the message was just sent and there
     * is a small wait time until it arrives.
     *
     * @param emailSnippets Snippets which contains a message search snippet that is needed
     * @param subjectLine   The subject line to search for
     * @return
     * @throws ExecutionException
     * @throws InterruptedException
     */
    protected Message GetAMessageFromEmailFolder(EmailSnippets emailSnippets, String subjectLine, String folderName)
            throws ExecutionException, InterruptedException {

        Message message = null;
        int tryCount = 0;

        //Continue trying to get the email while the email is not found
        //and the loop has tried less than MAX_POLL_REQUESTS times.
        do {
            List<Message> messages;
            messages = emailSnippets
                    .getMailboxMessagesByFolderNameSubject(
                            subjectLine
                            , folderName);
            if (messages.size() > 0) {
                message = messages.get(0);
            }
            tryCount++;
            Thread.sleep(THREAD_SLEEP_TIME);
            //Stay in loop while these conditions are true.
            //If either condition becomes false, break
        } while (message == null && tryCount < MAX_POLL_REQUESTS);

        return message;
    }

    //Deletes all messages with the given subject line from a named email folder
    protected void DeleteAMessageFromMailFolder(
            EmailSnippets emailSnippets
            , String subjectLine, String folderName)
            throws ExecutionException, InterruptedException {

        List<Message> messagesToDelete;
        int tryCount = 0;
        //Try to get the newly sent email from user's inbox at least once.
        //continue trying to get the email while the email is not found
        //and the loop has tried less than 50 times.
        do {

            messagesToDelete = emailSnippets
                    .getMailboxMessagesByFolderNameSubject(
                            subjectLine
                            , folderName);
            for (Message message : messagesToDelete) {
                //3. Delete the email using the ID
                emailSnippets.deleteMail(message.getId());
            }
            tryCount++;
            Thread.sleep(THREAD_SLEEP_TIME);
            //Stay in loop while these conditions are true.
            //If either condition becomes false, break
        } while (messagesToDelete.size() == 0 && tryCount < MAX_POLL_REQUESTS);

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
