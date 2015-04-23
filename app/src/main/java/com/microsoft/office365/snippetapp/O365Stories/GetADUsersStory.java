package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.directoryservices.User;
import com.microsoft.directoryservices.odata.DirectoryClient;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.O365ServicesManager;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.services.odata.impl.ADALDependencyResolver;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetADUsersStory extends BaseUserStory {

    @Override
    public String execute() {
        StringBuilder results = new StringBuilder();
        AuthenticationController
                .getInstance()
                .setResourceId(Constants.DIRECTORY_RESOURCE_ID);
//        ADALDependencyResolver dependencyResolver = (ADALDependencyResolver) AuthenticationController
//                .getInstance()
//                .getDependencyResolver();
//        dependencyResolver.setResourceId("https://graph.windows.net/");

//        DirectoryClient directoryClient=new DirectoryClient(Constants.DIRECTORY_RESOURCE_URL+AuthenticationController.TENANTID+"?api-version=1.5",dependencyResolver);
              UsersAndGroupsSnippets usersAndGroupsSnippets = new UsersAndGroupsSnippets(O365ServicesManager.getDirectoryClient());
        List<User> userList=null;
        try {
            userList = usersAndGroupsSnippets.getUsers();
            //Log.i("userTitle", userList.get(0).getdisplayName());
        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            return StoryResultFormatter.wrapResult("Get Active Directory users exception:",false);
        }

        //
        if (userList == null){
            //No users were found
            return StoryResultFormatter.wrapResult("Get Active Directory Users: No users found",true);
        }
        results.append("Get Active Directory Users: The following users were found:\n");
        for(User user : userList){
            results.append(user.getdisplayName())
                    .append("\n");
        }
        return StoryResultFormatter.wrapResult(results.toString(),true);
    }

    @Override
    public String getDescription() {
        return "Gets users from Active Directory";
    }
}
