package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.fileservices.Item;
import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

public class CreateOneDriveFolderStory extends BaseUserStory {
    //Unique names used for tracking and cleanup of items created by running snippets
    private static final String FOLDER_NAME = "O365SnippetFolder_create_";

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
            Item createdFolder = fileFolderSnippets.createO365Folder(folderName);

            //CLEANUP
            fileFolderSnippets.deleteO365Folder(folderName);

            return StoryResultFormatter.wrapResult(
                    "OneDrive create folder story: Folder created.", true);

        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("CreateOneDriveFolder", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Create OneDrive folder exception: "
                            + formattedException
                    , false
            );

        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("CreateOneDriveFolder", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Create OneDrive folder exception: "
                            + formattedException
                    , false
            );
        }
    }

    @Override
    public String getDescription() {
        return "Create new folder on OneDrive";
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
