package com.microsoft.office365.snippetapp.UserGroupStories;

import com.microsoft.directoryservices.User;
import com.microsoft.directoryservices.odata.DirectoryClient;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.O365ServicesManager;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetADUsersStory extends BaseUserStory {

    private static final String STORY_DESCRIPTION = "Gets users from Active Directory";

    @Override
    public String execute() {
        boolean isStoryComplete;
        StringBuilder resultMessage = new StringBuilder();

        AuthenticationController
                .getInstance()
                .setResourceId(Constants.DIRECTORY_RESOURCE_ID);

        DirectoryClient directoryClient = O365ServicesManager.getDirectoryClient();
        if (directoryClient == null)
            return StoryResultFormatter.wrapResult("Tenant ID was null", false);

        UsersAndGroupsSnippets usersAndGroupsSnippets = new UsersAndGroupsSnippets(directoryClient);

        try {
            //Get list of users
            List<User> userList = usersAndGroupsSnippets.getUsers();
            if (userList == null) {
                //No users were found
                resultMessage.append(STORY_DESCRIPTION);
                resultMessage.append(": No users found.");
            } else {
                resultMessage.append(STORY_DESCRIPTION);
                resultMessage.append(": The following users were found:\n");
                for (User user : userList) {
                    resultMessage.append(user.getdisplayName())
                            .append("\n");
                }
            }
        } catch (ExecutionException | InterruptedException e) {
            return FormatException(e, STORY_DESCRIPTION);
        }
        return StoryResultFormatter.wrapResult(resultMessage.toString(), true);
    }

    @Override
    public String getDescription() {
        return STORY_DESCRIPTION;
    }
}
