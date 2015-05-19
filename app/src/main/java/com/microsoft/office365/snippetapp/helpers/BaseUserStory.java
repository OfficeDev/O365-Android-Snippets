/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.helpers;


import android.content.res.AssetFileDescriptor;
import android.view.View;

import com.microsoft.fileservices.odata.SharePointClient;
import com.microsoft.office365.snippetapp.AndroidSnippetsApplication;
import com.microsoft.office365.snippetapp.Interfaces.OnUseCaseStatusChangedListener;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;

public abstract class BaseUserStory {

    private boolean mIsExecuting = false;
    private View mUpdateView;
    private OutlookClient mO365MailClient;
    private SharePointClient mO365MyFilesClient;
    private String mMailResourceId;
    private OnUseCaseStatusChangedListener mUseCaseStatusChangedListener;
    private String mFilesFoldersResourceId;
    boolean mGroupingFlag = false;

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


    public  boolean getGroupingFlag(){
        return mGroupingFlag;
    }

    public void setGroupingFlag(boolean groupingFlag){
        mGroupingFlag = groupingFlag;
    }
    public String getId() {
        return java.util.UUID.randomUUID().toString();
    }

    public String getStringResource(int resourceToGet) {
        return AndroidSnippetsApplication
                .getApplication()
                .getApplicationContext()
                .getString(resourceToGet);
    }

    public byte[] getDrawableResource(int resourceToGet) {

        //Get the photo from the resource/drawable folder as a raw image
        final AssetFileDescriptor raw = AndroidSnippetsApplication
                .getApplication()
                .getApplicationContext()
                .getResources()
                .openRawResourceFd(resourceToGet);

        //Load raw image into a buffer
        final ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        try {
            final FileInputStream is = raw.createInputStream();
            int nRead;

            //Read 16kb at a time
            final byte[] data = new byte[16384];

            while ((nRead = is.read(data, 0, data.length)) != -1) {
                buffer.write(data, 0, nRead);
            }

            buffer.flush();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return buffer.toByteArray();

    }

    public String FormatExceptionMessage(Exception exception)
    {
        String formattedException = APIErrorMessageHelper.getErrorMessage(exception.getMessage());
        return StoryResultFormatter.wrapResult(
                "Forward email message story: " + formattedException, false);

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
