/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;


import android.view.View;

import com.microsoft.fileservices.odata.SharePointClient;
import com.microsoft.office365.snippetapp.AndroidSnippetsApplication;
import com.microsoft.office365.snippetapp.Interfaces.OnUseCaseStatusChangedListener;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.outlookservices.Message;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.util.List;
import java.util.concurrent.ExecutionException;

public abstract class BaseUserStory {

    private boolean mIsExecuting = false;
    private View mUpdateView;
    private OutlookClient mO365MailClient;
    private SharePointClient mO365MyFilesClient;
    private String mMailResourceId;
    private OnUseCaseStatusChangedListener mUseCaseStatusChangedListener;
    private String mFilesFoldersResourceId;
    private static final int MAX_POLL_REQUESTS = 20;
    EmailSnippets mEmailSnippets;

    public String getFilesFoldersResourceId() {
        return mFilesFoldersResourceId;
    }

    public void setFilesFoldersResourceId(String filesFoldersResourceId) {
        this.mFilesFoldersResourceId = filesFoldersResourceId;
    }

    public void setUseCaseStatusChangedListener(OnUseCaseStatusChangedListener listener) {
        mUseCaseStatusChangedListener = listener;
    }

    public abstract String execute();

    public abstract String getDescription();

    public String getId() {
        return java.util.UUID.randomUUID().toString();
    }

    public String getStringResource(int resourceToGet)
    {
        return AndroidSnippetsApplication
                .getApplication()
                .getApplicationContext()
                .getString(resourceToGet);
    }
    //Gets messages with the given subject line from the user's inbox
    protected Message GetAMessageFromInBox(String subjectLine) throws ExecutionException, InterruptedException{
        //Get the new message
        Message messageToAttach = null;
        int tryCount = 0;

        //Try to get the newly sent email from user's inbox at least once.
        //continue trying to get the email while the email is not found
        //and the loop has tried less than 50 times.
        do {
            List<Message> messages = null;
            messages = mEmailSnippets
                    .GetMailboxMessagesByFolderName_Subject(
                            subjectLine
                            , getStringResource(R.string.Email_Folder_Inbox));
            if (messages.size() > 0) {
                messageToAttach = messages.get(0);
            }
            tryCount++;
            //Stay in loop while these conditions are true.
            //If either condition becomes false, break
        } while (messageToAttach != null && tryCount < MAX_POLL_REQUESTS);

        return messageToAttach;
    }

    //Deletes all messages with the given subject line from a named email folder
    protected void DeleteAMessageFromMailFolder(String subjectLine, String folderName) throws ExecutionException, InterruptedException{
        List<Message> messagesToDelete = null;
        int tryCount = 0;
        //Try to get the newly sent email from user's inbox at least once.
        //continue trying to get the email while the email is not found
        //and the loop has tried less than 50 times.
        do {

            messagesToDelete = mEmailSnippets
                    .GetMailboxMessagesByFolderName_Subject(
                            subjectLine
                            , folderName);
            for (Message message : messagesToDelete) {
                //3. Delete the email using the ID
                mEmailSnippets.deleteMail(message.getId());
            }
            tryCount++;
            //Stay in loop while these conditions are true.
            //If either condition becomes false, break
        } while (messagesToDelete.size() > 0 && tryCount < MAX_POLL_REQUESTS);

    }

    public View getUIResultView() {
        return mUpdateView;
    }

    public void setUIResultView(View view) {
        mUpdateView = view;
    }

    public String getO365MailResourceId() {
        return mMailResourceId;
    }

    public void setO365MailResourceId(String resourceId) {
        mMailResourceId = resourceId;
    }

    public OutlookClient getO365MailClient() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        this.getO365MailResourceId());

        return mO365MailClient;
    }

    public void setO365MailClient(OutlookClient client) {

        mO365MailClient = client;
    }

    public SharePointClient getO365MyFilesClient() {
        return mO365MyFilesClient;
    }

    public void setO365MyFilesClient(SharePointClient client) {
        mO365MyFilesClient = client;
    }

    private void notifyStatusChange() {
        if (null != mUseCaseStatusChangedListener) {
            mUseCaseStatusChangedListener.onUseCaseStatusChanged();
        }
    }

    public final void onPreExecute() {
        mIsExecuting = true;
        notifyStatusChange();
    }

    public final void onPostExecute() {
        mIsExecuting = false;
        notifyStatusChange();
    }

    public final boolean isExecuting() {
        return mIsExecuting;
    }

    @Override
    public String toString() {
        return this.getDescription();
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
