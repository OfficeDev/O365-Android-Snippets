/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.fileservices.Item;
import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.List;
import java.util.concurrent.ExecutionException;

/**
 * Created by Microsoft on 3/19/15.
 */
public class GetFilesAndFoldersStory extends BaseUserStory {
    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getFilesFoldersResourceId());

        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(
                getO365MyFilesClient());
        try {
            List<Item> items = fileFolderSnippets.getFilesAndFolders();
            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append("Gets items: "
                    + items.size()
                    + " items returned");
            sb.append("\n");
            for (Item item : items) {
                sb.append("\t\t");
                sb.append(item.gettype() + ": " + item.getname());
                sb.append("\n");
            }
            return StoryResultFormatter.wrapResult(sb.toString(), true);

        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Get Files/folders", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Get Files and folders exception: "
                            + formattedException, false
            );
        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Get Files/folders", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Get Files and folders exception: "
                            + formattedException, false
            );
        }

    }

    @Override
    public String getDescription() {
        return "Gets files and folders from user's OneDrive";
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
