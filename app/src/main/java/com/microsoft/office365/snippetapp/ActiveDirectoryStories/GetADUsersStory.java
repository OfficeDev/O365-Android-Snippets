package com.microsoft.office365.snippetapp.ActiveDirectoryStories;

import android.util.Log;

import com.microsoft.directoryservices.User;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.O365ServicesManager;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetADUsersStory extends BaseUserStory {

    @Override
    public String execute() {
        boolean isStoryComplete;
        StringBuilder resultMessage = new StringBuilder();

        AuthenticationController
                .getInstance()
                .setResourceId(Constants.DIRECTORY_RESOURCE_ID);
        UsersAndGroupsSnippets usersAndGroupsSnippets = new UsersAndGroupsSnippets(O365ServicesManager.getDirectoryClient());

        try {
            //Get list of users
            List<User> userList = usersAndGroupsSnippets.getUsers();
            if (userList == null) {
                //No users were found
                resultMessage.append("Get Active Directory Users: No users found.");
            } else {
                resultMessage.append("Get Active Directory Users: The following users were found:\n");
                for (User user : userList) {
                    resultMessage.append(user.getdisplayName())
                            .append("\n");
                }
            }
            isStoryComplete = true;
        } catch (ExecutionException | InterruptedException e) {
            isStoryComplete = false;
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("GetADUsers", formattedException);
            resultMessage = new StringBuilder();
            resultMessage.append("Get Active Directory users exception: ")
                    .append(formattedException);
        }
        return StoryResultFormatter.wrapResult(resultMessage.toString(), isStoryComplete);
    }

    @Override
    public String getDescription() {
        return "Gets users from Active Directory";
    }
}
