package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.fileservices.Item;
import com.microsoft.office365.snippetapp.Snippets.FileFolderSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

public class DeleteOneDriveFolderStory extends BaseUserStory {
    //Unique names used for tracking and cleanup of items created by running snippets
    private static final String FOLDER_NAME = "O365SnippetFolder_delete_";

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(getFilesFoldersResourceId());
        FileFolderSnippets fileFolderSnippets = new FileFolderSnippets(
                getO365MyFilesClient());

        try {
            String folderName = FOLDER_NAME + java
                    .util
                    .UUID
                    .randomUUID()
                    .toString();

            Item createdFolder = fileFolderSnippets.createO365Folder(folderName);
            fileFolderSnippets.deleteO365Folder(folderName);

            return StoryResultFormatter.wrapResult(
                    "OneDrive delete folder story: Folder deleted.", true);

        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("DeleteOneDriveFolder", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Decline OneDrive folder exception: "
                            + formattedException
                    , false
            );

        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("DeleteOneDriveFolder", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Delete OneDrive folder exception: "
                            + formattedException
                    , false
            );
        }
    }

    @Override
    public String getDescription() {
        return "Delete folder from OneDrive";
    }
}
