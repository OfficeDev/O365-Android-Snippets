/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import com.google.common.base.Charsets;
import com.microsoft.fileservices.File;
import com.microsoft.fileservices.Item;
import com.microsoft.services.odata.Constants;
import com.microsoft.sharepointservices.odata.SharePointClient;

import java.util.List;
import java.util.concurrent.ExecutionException;

/**
 * Created by Microsoft on 3/19/15.
 */
public class FileFolderSnippets {
    SharePointClient mSharePointClient;

    public FileFolderSnippets(SharePointClient sharePointClient) {
        mSharePointClient = sharePointClient;
    }

    public List<Item> getFilesAndFolders()
            throws ExecutionException
            , InterruptedException {
        List<Item> filesAndFolders = mSharePointClient.getfiles().read().get();
        return filesAndFolders;
    }

    public String getFileFromServerById(String fileId) throws ExecutionException, InterruptedException {
        File retrievedFile = mSharePointClient.getfiles()
                .getById(fileId)
                .asFile()
                .read().get();

        return retrievedFile.getid();
    }

    public String getFileFromServerByName(String fileName) throws ExecutionException, InterruptedException {
        String itemID = "";
        List<Item> filesAndFolders = mSharePointClient.getfiles().read().get();

        for (Item item : filesAndFolders) {
            if (item.getname().equals(fileName)) {
                itemID = item.getid();
            }
        }
        return itemID;
    }

    public byte[] getFileContentsFromServer(String fileId)
            throws ExecutionException
            , InterruptedException {
        byte[] fileContents = mSharePointClient.getfiles()
                .getById(fileId)
                .asFile()
                .getContent().get();
        return fileContents;
    }

    public void deleteFileFromServer(String fileId)
            throws ExecutionException
            , InterruptedException {
        mSharePointClient.getfiles()
                .getById(fileId)
                .addHeader("If-Match", "*")
                .delete()
                .get();

    }

    public String postNewFileToServer(
            String fileName
            , byte[] fileContents)
            throws ExecutionException
            , InterruptedException {
        Item newFile = new Item();
        newFile.settype("File");
        newFile.setname(fileName);

        newFile = mSharePointClient
                .getfiles()
                .add(newFile)
                .get();

        mSharePointClient.getfiles()
                .getById(newFile.getid())
                .asFile()
                .putContent(fileContents)
                .get();
        return newFile.getid();
    }

    public void postUpdatedFileToServer(
            String fileId
            , String updatedContents)
            throws ExecutionException
            , InterruptedException {
        mSharePointClient.getfiles()
                .getById(fileId)
                .asFile()
                .putContent(
                        updatedContents.getBytes(Charsets.UTF_8)).get();
    }

    public Item createO365Folder(String fullPath) throws ExecutionException, InterruptedException {
        Item folder = new Item();

        folder.settype("Folder");
        folder.setname(fullPath);
        Item createdFolder = mSharePointClient.getfiles().add(folder).get();
        return createdFolder;
    }

    public void deleteO365Folder(String fullPath) throws ExecutionException, InterruptedException {
        //Find ID of the path
        Item folder = mSharePointClient
                .getfiles()
                .getOperations()
                .getByPath(fullPath)
                .get();

        //Use ID to delete the folder
        mSharePointClient
                .getfiles()
                .getById(folder.getid())
                .addHeader(Constants.IF_MATCH_HEADER, "*")
                .delete()
                .get();
    }
}
// *********************************************************
//
// O365-Android-Snippet, https://github.com/OfficeDev/O365-Android-Snippet
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
