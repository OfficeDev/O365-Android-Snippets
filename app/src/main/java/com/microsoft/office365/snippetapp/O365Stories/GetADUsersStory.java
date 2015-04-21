package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.directoryservices.User;
import com.microsoft.directoryservices.odata.DirectoryClient;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.services.odata.impl.ADALDependencyResolver;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetADUsersStory extends BaseUserStory {

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId("https://graph.windows.net/");
        ADALDependencyResolver dependencyResolver = (ADALDependencyResolver) AuthenticationController
                .getInstance()
                .getDependencyResolver();
        DirectoryClient directoryClient=new DirectoryClient("https://graph.windows.net/"+AuthenticationController.TENANTID,dependencyResolver);

        try {
            List<User> userList = directoryClient.getusers().read().get();
            Log.i("userTitle", userList.get(0).getdisplayName());
        } catch (ExecutionException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        return "well";

    }

    @Override
    public String getDescription() {
        return "Gets users from Active Directory";
    }
}
