/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import com.microsoft.directoryservices.Group;
import com.microsoft.directoryservices.TenantDetail;
import com.microsoft.directoryservices.User;
import com.microsoft.directoryservices.odata.DirectoryClient;

import java.util.List;
import java.util.concurrent.ExecutionException;


public class UsersAndGroupsSnippets {

    DirectoryClient mDirectoryClient;

    public UsersAndGroupsSnippets(DirectoryClient directoryClient) {
        mDirectoryClient = directoryClient;
    }

    /**
     * Return a list of users from Active Directory.
     *
     * @return List. A list of the com.microsoft.directoryservices.User objects.
     * @version 1.0
     */
    public List<User> getUsers() throws ExecutionException, InterruptedException {
        return mDirectoryClient.getusers().read().get();
    }

    /**
     * Return tenant details from Active Directory.
     * *
     *
     * @return TenantDetail. The com.microsoft.directoryservices.TenantDetail object for first tenant found.
     * @version 1.0
     */
    public TenantDetail getTenantDetails() throws ExecutionException, InterruptedException {
        List<TenantDetail> tenants = mDirectoryClient.gettenantDetails().read().get();
        return tenants.get(0);
    }

    /**
     * Return a list of groups from Active Directory.
     *
     * @return List. A list of the com.microsoft.directoryservices.Group objects.
     * @version 1.0
     */
    public List<Group> getGroups() throws ExecutionException, InterruptedException {
        return mDirectoryClient.getgroups().read().get();
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


