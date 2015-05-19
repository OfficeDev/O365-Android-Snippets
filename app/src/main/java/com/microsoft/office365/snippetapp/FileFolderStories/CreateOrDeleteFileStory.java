/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.FileFolderStories;

import com.google.common.base.Charsets;
import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryAction;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

//This story handles both of the following stories that appear in the UI list
// based on strings passed in the constructor...
//- Create a file (which is then deleted for cleanup)
//- Delete a file (which is created first and then deleted)
public class CreateOrDeleteFileStory extends BaseUserStory {
    private final String CREATE_DESCRIPTION = "Create a file on server";
    private final String CREATE_TAG = "CreateFile";
    private final String CREATE_SUCCESS = "Create a file on server";
    private final String CREATE_ERROR = "Create a file on server exception: ";

    private final String DELETE_DESCRIPTION = "Delete a file on server";
    private final String DELETE_TAG = "DeleteFile";
    private final String DELETE_SUCCESS = "Delete a file on server";
    private final String DELETE_ERROR = "Delete a file on server exception: ";

    private String mDescription;
    private String mLogTag;
    private String mSuccessDescription;
    private String mErrorDescription;

    public CreateOrDeleteFileStory(StoryAction action) {
        switch (action) {
            case CREATE: {
                mDescription = CREATE_DESCRIPTION;
                mLogTag = CREATE_TAG;
                mSuccessDescription = CREATE_SUCCESS;
                mErrorDescription = CREATE_ERROR;
                break;
            }
            case DELETE: {
                mDescription = DELETE_DESCRIPTION;
                mLogTag = DELETE_TAG;
                mSuccessDescription = DELETE_SUCCESS;
                mErrorDescription = DELETE_ERROR;
                break;
            }
        }
    }

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getFilesFoldersResourceId());
        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(
                getO365MyFilesClient());

        try {

            String fileContents = "Test create and delete file";
            //Create file
            String newFileId = fileFolderSnippets
                    .postNewFileToServer(
                            "test_Create_Delete_"
                                    + java
                                    .util
                                    .UUID
                                    .randomUUID()
                                    .toString()
                                    + ".txt"
                            , fileContents.getBytes(Charsets.UTF_8));

            //Delete file
            fileFolderSnippets.deleteFileFromServer(newFileId);
            return StoryResultFormatter.wrapResult(mSuccessDescription, true);
        } catch (ExecutionException | InterruptedException e) {
            return BaseExceptionFormatter(e, mDescription);
        }
    }

    @Override
    public String getDescription() {
        return mDescription;
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
