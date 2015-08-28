/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import com.google.common.base.Charsets;
import com.microsoft.services.files.Item;
import com.microsoft.services.files.fetchers.FilesClient;
import com.microsoft.services.odata.Constants;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class FileFolderSnippets {
    FilesClient mFilesClient;

    public FileFolderSnippets(FilesClient sharePointClient) {
        mFilesClient = sharePointClient;
    }

    /**
     * Gets the folders and files at the user's OneDrive for business root
     *
     * @return List. A list of the com.microsoft.fileservices.Item objects
     */
    public List<Item> getFilesAndFolders()
            throws ExecutionException
            , InterruptedException {
        return mFilesClient
               .getFiles()
                .select("name,type")
                .read()
                .get();
    }

    /**
     * Gets the ID of a file on user's OnDrive for Business by file name
     *
     * @param fileName The name of the file whose Id is to be returned
     * @return String. The Id of the retrieved file
     */
    public String getFileFromServerByName(String fileName) throws ExecutionException, InterruptedException {
        String itemID = "";
        List<Item> filesAndFolders = mFilesClient
                .getFiles()
                .select("Name")
                .read()
                .get();

        for (Item item : filesAndFolders) {
            if (item.getName().equals(fileName)) {
                itemID = item.getId();
            }
        }
        return itemID;
    }

    /**
     * Gets the contents of a file on user's OnDrive for Business by ID
     *
     * @param fileId The Id of the file whose contents are to be returned
     * @return Byte[]. The contents of the file as a byte array
     */
    public byte[] getFileContentsFromServer(String fileId)
            throws ExecutionException, InterruptedException {
        byte[] fileContents = mFilesClient.getFiles()
                .getById(fileId)
                .asFile()
                .getContent().get();
        return fileContents;
    }

    /**
     * Deletes a file on user's OnDrive for Business by ID
     *
     * @param fileId The Id of the file whose contents are to be returned
     */
    public void deleteFileFromServer(String fileId)
            throws ExecutionException, InterruptedException {
        mFilesClient.getFiles()
                .getById(fileId)
                .addHeader("If-Match", "*")
                .delete()
                .get();

    }

    /**
     * Uploads a file to the root folder of a user's OneDrive for Business drive
     *
     * @param fileName     The name of the file to be uploaded
     * @param fileContents Byte[]. The contents of the file as a byte array
     */
    public String postNewFileToServer(
            String fileName
            , byte[] fileContents)
            throws ExecutionException
            , InterruptedException {
        Item newFile = new Item();
        newFile.setType("File");
        newFile.setName(fileName);

        newFile = mFilesClient
                .getFiles()
                .select("ID")
                .add(newFile)
                .get();

        mFilesClient.getFiles()
                .getById(newFile.getId())
                .asFile()
                .putContent(fileContents)
                .get();
        return newFile.getId();
    }

    /**
     * Uploads a an update to a file to the root folder of a user's OneDrive for Business drive
     *
     * @param fileId          The id of the file to be uploaded
     * @param updatedContents The contents of the file as a string
     */
    public void postUpdatedFileToServer(
            String fileId
            , String updatedContents)
            throws ExecutionException
            , InterruptedException {
        mFilesClient.getFiles()
                .getById(fileId)
                .asFile()
                .putContent(
                        updatedContents.getBytes(Charsets.UTF_8)).get();
    }

    /**
     * Creates a new folder in the root of the user's OneDrive for Business drive
     *
     * @param fullPath The path of the folder to be created
     * @return Item  The created folder
     */
    public Item createO365Folder(String fullPath) throws ExecutionException, InterruptedException {
        Item folder = new Item();

        folder.setType("Folder");
        folder.setName(fullPath);
        Item createdFolder = mFilesClient
                .getFiles()
                .select("ID")
                .add(folder)
                .get();
        return createdFolder;
    }

    /**
     * Deletes a folder from the user's OneDrive for Business drive
     *
     * @param fullPath The path of the folder to be removed
     */
    public void deleteO365Folder(String fullPath) throws ExecutionException, InterruptedException {
        //Find ID of the path
        Item folder = mFilesClient
                .getFiles()
                .getOperations()
                .getByPath(fullPath)
                .get();

        //Use ID to delete the folder
        mFilesClient
                .getFiles()
                .getById(folder.getId())
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
