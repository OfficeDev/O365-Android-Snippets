/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.FileFolderStories;

import com.google.common.base.Charsets;
import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

public class UpdateFileContentsOnServerStory extends BaseUserStory {

    private static final String STORY_DESCRIPTION = "Update file contents on user's OneDrive";

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getFilesFoldersResourceId());
        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(
                getO365MyFilesClient());
        try {

            String fileContents = "Test create file";
            String newFileId = fileFolderSnippets
                    .postNewFileToServer(
                            "test_Update_"
                                    + java
                                    .util
                                    .UUID
                                    .randomUUID()
                                    .toString()
                                    + ".txt"
                            , fileContents.getBytes(Charsets.UTF_8));


            String updatedFileContents = fileContents + " updated";
            fileFolderSnippets.postUpdatedFileToServer(newFileId, updatedFileContents);
            byte[] fileContentsBytes = fileFolderSnippets.getFileContentsFromServer(newFileId);
            fileFolderSnippets.deleteFileFromServer(newFileId);

            if (fileContentsBytes.length == updatedFileContents.length()) {
                return StoryResultFormatter.wrapResult(STORY_DESCRIPTION, true);
            } else {
                return StoryResultFormatter.wrapResult(STORY_DESCRIPTION, false);
            }

        } catch (ExecutionException | InterruptedException e) {
            return BaseExceptionFormatter(e, STORY_DESCRIPTION);
        }
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
