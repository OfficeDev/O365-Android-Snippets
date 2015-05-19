/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.FileFolderStories;

import android.os.Environment;
import android.util.Log;

import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.ExecutionException;

public class DownloadFileStory extends BaseUserStory {
    private static final String DOWNLOAD_DOC_PATH = "O365Snippets";
    private static final String DOWNLOAD_DOC_FILENAME = "testdownload.txt";
    private static final String FILE_CONTENTS = "Test download file contents";
    private boolean mFolderCreated = false;

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getFilesFoldersResourceId());
        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(
                getO365MyFilesClient());
        try {
            String fileContents = FILE_CONTENTS;
            String itemId;

            //Remove test text file from server if left by previous execute()
            itemId = fileFolderSnippets.getFileFromServerByName(DOWNLOAD_DOC_FILENAME);
            if (itemId.length() != 0) {
                //Remove remote file instance
                fileFolderSnippets.deleteFileFromServer(itemId);
            }
            String newFileId = fileFolderSnippets
                    .postNewFileToServer(
                            DOWNLOAD_DOC_FILENAME
                            , fileContents.getBytes());


            byte[] fileContentsBytes = fileFolderSnippets.getFileContentsFromServer(newFileId);

            //Remove remote file instance
            fileFolderSnippets.deleteFileFromServer(newFileId);

            //Save file locally
            if (saveFileInExternalStorage(DOWNLOAD_DOC_FILENAME, fileContentsBytes)) {

                //Verify local file exists and then remove it
                if (verifyFileInExternalStorage(DOWNLOAD_DOC_FILENAME))
                    return StoryResultFormatter.wrapResult("Download file from server", true);
                else
                    return StoryResultFormatter.wrapResult("Download file from server", false);

            } else
                return StoryResultFormatter.wrapResult("Download file from server", false);
        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update file on server", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update file contents on server exception: "
                            + formattedException, false
            );
        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update file on server", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update file contents on server exception: "
                            + formattedException, false
            );
        }
    }

    @Override
    public String getDescription() {
        return "Download a file from MyFiles";
    }


    //Verifies that the downloaded file is saved in external storage and
    //then deletes the downloaded file and the folder that it is saved in
    private boolean verifyFileInExternalStorage(String fileName) {
        boolean fileExistsFlag = false;
        File fileDocDir = verifyDownloadFolder();
        if (fileDocDir == null) return fileExistsFlag;

        final File downloadedFile = new File(fileDocDir.getPath(), fileName);
        if (downloadedFile.exists()) {
            fileExistsFlag = deleteDownloadedFileandFolder(downloadedFile, fileDocDir);
        }
        return fileExistsFlag;
    }

    //Delete the passed file and folder from external storage
    private boolean deleteDownloadedFileandFolder(File downloadedFile, File fileDocDir) {
        if (downloadedFile.exists()) {
            downloadedFile.delete();
            if (mFolderCreated) {
                fileDocDir.delete();
            }
            return true;
        }
        return false;
    }

    private boolean saveFileInExternalStorage(String fileName, final byte[] fileContents) {
        File fileDocDir = verifyDownloadFolder();
        if (fileDocDir == null) return false;

        if (!fileDocDir.exists() && (!fileDocDir.mkdirs())) {
            Log.e("IO error", "Directory not created");
            return false;
        } else {
            final File downloadedFile = new File(fileDocDir.getPath(), fileName);
            if (downloadedFile.exists()) downloadedFile.delete();
            writeFileOut(downloadedFile, fileContents);
            return true;
        }
    }

    private void writeFileOut(File downloadedFile, byte[] filecontent) {
        try {
            FileOutputStream outputStream = new FileOutputStream(downloadedFile);
            try {
                outputStream.write(filecontent);
            } finally {
                outputStream.close();
            }
        } catch (IOException io) {
            Log.e(
                    "File IO ERROR"
                    , "The file was not saved "
                            + io.getMessage()
            );
        }
    }

    private File verifyDownloadFolder() {
        File fileDocDir = null;
        String externalStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(externalStorageState)) {
            fileDocDir = new File(
                    Environment
                            .getExternalStorageDirectory().getPath(), DOWNLOAD_DOC_PATH
            );

            if (!fileDocDir.exists()) {
                mFolderCreated = true;
                fileDocDir.mkdirs();
            }
        }
        return fileDocDir;
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
