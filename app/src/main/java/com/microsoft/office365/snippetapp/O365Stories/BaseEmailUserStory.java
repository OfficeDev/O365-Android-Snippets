/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
package com.microsoft.office365.snippetapp.O365Stories;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.outlookservices.Message;

import java.util.List;
import java.util.concurrent.ExecutionException;

public abstract class BaseEmailUserStory extends BaseUserStory {

    private static final int MAX_POLL_REQUESTS = 20;

    public abstract String execute();

    public abstract String getDescription();

    //Gets messages with the given subject line from the user's inbox
    protected Message GetAMessageFromInBox(EmailSnippets emailSnippets, String subjectLine) throws ExecutionException, InterruptedException {
        //Get the new message
        Message messageToAttach = null;
        int tryCount = 0;

        //Try to get the newly sent email from user's inbox at least once.
        //continue trying to get the email while the email is not found
        //and the loop has tried less than 50 times.
        do {
            List<Message> messages = null;
            messages = emailSnippets
                    .GetMailboxMessagesByFolderName_Subject(
                            subjectLine
                            , getStringResource(R.string.Email_Folder_Inbox));
            if (messages.size() > 0) {
                messageToAttach = messages.get(0);
            }
            tryCount++;
            //Stay in loop while these conditions are true.
            //If either condition becomes false, break
        } while (messageToAttach == null && tryCount < MAX_POLL_REQUESTS);

        return messageToAttach;
    }

    //Deletes all messages with the given subject line from a named email folder
    protected void DeleteAMessageFromMailFolder(
            EmailSnippets emailSnippets
            , String subjectLine, String folderName) throws ExecutionException, InterruptedException {
        List<Message> messagesToDelete = null;
        int tryCount = 0;
        //Try to get the newly sent email from user's inbox at least once.
        //continue trying to get the email while the email is not found
        //and the loop has tried less than 50 times.
        do {

            messagesToDelete = emailSnippets
                    .GetMailboxMessagesByFolderName_Subject(
                            subjectLine
                            , folderName);
            for (Message message : messagesToDelete) {
                //3. Delete the email using the ID
                emailSnippets.deleteMail(message.getId());
            }
            tryCount++;
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