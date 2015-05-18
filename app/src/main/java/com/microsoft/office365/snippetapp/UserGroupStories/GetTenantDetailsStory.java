/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.UserGroupStories;

import com.microsoft.directoryservices.TenantDetail;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.Snippets.UsersAndGroupsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.O365ServicesManager;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

public class GetTenantDetailsStory extends BaseUserStory {
    @Override
    public String execute() {
        StringBuilder results = new StringBuilder();
        AuthenticationController
                .getInstance()
                .setResourceId(Constants.DIRECTORY_RESOURCE_ID);

        UsersAndGroupsSnippets usersAndGroupsSnippets = new UsersAndGroupsSnippets(O365ServicesManager.getDirectoryClient());
        TenantDetail tenant;
        try {
            tenant = usersAndGroupsSnippets.getTenantDetails();
        } catch (ExecutionException | InterruptedException e) {
            e.printStackTrace();
            return StoryResultFormatter.wrapResult("Get tenant detail exception:", false);
        }

        if (tenant == null) {
            //No tenants were found
            return StoryResultFormatter.wrapResult("Get tenant detail: No tenant found", true);
        }
        results.append("Get Active Directory Users: The following tenant was found:\n");
        results.append(tenant.getdisplayName())
                .append("\n");
        return StoryResultFormatter.wrapResult(results.toString(), true);
    }

    @Override
    public String getDescription() {
        return "Gets the tenant details from Active Directory";
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
