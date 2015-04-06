/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;


import android.view.View;

import com.microsoft.office365.snippetapp.Interfaces.OnUseCaseStatusChangedListener;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.outlookservices.odata.OutlookClient;
import com.microsoft.fileservices.odata.SharePointClient;

public abstract class BaseUserStory {

    private boolean mIsExecuting = false;
    private View mUpdateView;
    private OutlookClient mO365MailClient;
    private SharePointClient mO365MyFilesClient;
    private String mMailResourceId;
    private OnUseCaseStatusChangedListener mUseCaseStatusChangedListener;

    public String getFilesFoldersResourceId() {
        return mFilesFoldersResourceId;
    }

    public void setFilesFoldersResourceId(String filesFoldersResourceId) {
        this.mFilesFoldersResourceId = filesFoldersResourceId;
    }

    public void setUseCaseStatusChangedListener(OnUseCaseStatusChangedListener listener) {
        mUseCaseStatusChangedListener = listener;
    }

    private String mFilesFoldersResourceId;

    public abstract String execute();

    public abstract String getDescription();

    public String getId() {
        return java.util.UUID.randomUUID().toString();
    }


    public void setUIResultView(View view) {
        mUpdateView = view;
    }

    public View getUIResultView() {
        return mUpdateView;
    }

    public void setO365MailClient(OutlookClient client) {

        mO365MailClient = client;
    }

    public void setO365MyFilesClient(SharePointClient client) {
        mO365MyFilesClient = client;
    }

    public void setO365MailResourceId(String resourceId) {
        mMailResourceId = resourceId;
    }

    public String getO365MailResourceId() {
        return mMailResourceId;
    }

    public OutlookClient getO365MailClient() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        this.getO365MailResourceId());
        return mO365MailClient;
    }

    public SharePointClient getO365MyFilesClient() {
        return mO365MyFilesClient;
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
