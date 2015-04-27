package com.microsoft.office365.snippetapp.O365Stories;

import com.microsoft.directoryservices.User;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.O365ServicesManager;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetADUsersStory extends BaseUserStory {

    @Override
    public String execute() {
        StringBuilder results = new StringBuilder();
        AuthenticationController
                .getInstance()
                .setResourceId(Constants.DIRECTORY_RESOURCE_ID);

        UsersAndGroupsSnippets usersAndGroupsSnippets = new UsersAndGroupsSnippets(O365ServicesManager.getDirectoryClient());
        List<User> userList = null;
        try {
            userList = usersAndGroupsSnippets.getUsers();
        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            return StoryResultFormatter.wrapResult("Get Active Directory users exception:", false);
        }

        //
        if (userList == null) {
            //No users were found
            return StoryResultFormatter.wrapResult("Get Active Directory Users: No users found", true);
        }
        results.append("Get Active Directory Users: The following users were found:\n");
        for (User user : userList) {
            results.append(user.getdisplayName())
                    .append("\n");
        }
        return StoryResultFormatter.wrapResult(results.toString(), true);
    }

    @Override
    public String getDescription() {
        return "Gets users from Active Directory";
    }
}
