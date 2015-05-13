/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryAction;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

//This story handles both of the following stories that appear in the UI list
// based on strings passed in the constructor...
//- Create a OneDrive Folder (which is then deleted for cleanup)
//- Delete a OneDrive Folder (which is created first and then deleted)
public class CreateOrDeleteOneDriveFolder extends BaseUserStory {
    //Unique names used for tracking and cleanup of items created by running snippets
    private static final String FOLDER_NAME = "O365SnippetFolder_";
    private final String CREATE_DESCRIPTION = "Create new folder on OneDrive";
    private final String CREATE_TAG = "CreateOneDriveFolder";
    private final String CREATE_SUCCESS = "OneDrive create folder story: Folder created.";
    private final String CREATE_ERROR = "Create OneDrive folder exception: ";
    private final String DELETE_DESCRIPTION = "Delete folder from OneDrive";
    private final String DELETE_TAG = "DeleteOneDriveFolder";
    private final String DELETE_SUCCESS = "OneDrive delete folder story: Folder deleted.";
    private final String DELETE_ERROR = "Decline OneDrive folder exception: ";
    private String mDescription;
    private String mLogTag;
    private String mSuccessDescription;
    private String mErrorDescription;

    public CreateOrDeleteOneDriveFolder(StoryAction action) {
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
        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(getO365MyFilesClient());

        try {
            String folderName = FOLDER_NAME + java
                    .util
                    .UUID
                    .randomUUID()
                    .toString();

            //Create folder
            fileFolderSnippets.createO365Folder(folderName);

            //Delete folder
            fileFolderSnippets.deleteO365Folder(folderName);

            return StoryResultFormatter.wrapResult(
                    mSuccessDescription, true);

        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e(mLogTag, formattedException);
            return StoryResultFormatter.wrapResult(mErrorDescription + formattedException
                    , false
            );
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
