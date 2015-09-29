/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;


import android.util.Log;

import com.microsoft.directoryservices.Group;
import com.microsoft.directoryservices.TenantDetail;
import com.microsoft.directoryservices.User;
import com.microsoft.directoryservices.odata.DirectoryClient;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.helpers.RetroFitHelpers;
import com.microsoft.office365.snippetapp.SnippetApp;
import java.util.List;
import java.util.concurrent.ExecutionException;

import retrofit.RetrofitError;
import retrofit.client.Response;


public class UsersAndGroupsSnippets  {

    DirectoryClient mDirectoryClient;
    UsersAndGroupsService mUsersAndGroupsService;
    SnippetApp mSnippetApp;

    public UsersAndGroupsSnippets(DirectoryClient directoryClient) {
        mDirectoryClient = directoryClient;
        mSnippetApp = SnippetApp.getApplication();
    }

    public UsersAndGroupsSnippets(String  AccessToken)
    {
        RetroFitHelpers retroFitHelpers = new RetroFitHelpers(AccessToken);
        mUsersAndGroupsService =  retroFitHelpers.getRestAdapter().create(UsersAndGroupsService.class);
        //Create UsersAndGroupsService...
    }

    private retrofit.Callback<Envelope<UserValue>> getUsersCallback() {
        return new retrofit.Callback<Envelope<UserValue>>() {
            @Override
            public void success(Envelope<UserValue> users, Response response) {

                Log.i("Users returned ", "Multiple users were returned: " + users.value.length);

            }

            @Override
            public void failure(RetrofitError error) {
                // TODO show an error and disable the run button
            }
        };
    }
    /**
     * Return a list of users from Active Directory, sorted by display name..
     *
     * @return List. A list of the com.microsoft.directoryservices.User objects.
     */
    public List<User> getUsers() throws ExecutionException, InterruptedException {

        mUsersAndGroupsService.getUsers(
                null,
                null,
                null,
                null,
                null,
                mSnippetApp.
                        getString(
                                R.string.appJSON) ,
                getUsersCallback());
//        return mDirectoryClient
//                .getusers()
//                .orderBy("displayName")
//                .read()
//                .get();
        return null;
    }

    /**
     * Return tenant details from Active Directory.
     * *
     *
     * @return TenantDetail. The com.microsoft.directoryservices.TenantDetail object for first tenant found.
     */
    public TenantDetail getTenantDetails() throws ExecutionException, InterruptedException {
        List<TenantDetail> tenants = mDirectoryClient
                .gettenantDetails()
                .read()
                .get();
        return tenants.get(0);
    }

    /**
     * Return a list of groups from Active Directory.
     *
     * @return List. A list of the com.microsoft.directoryservices.Group objects.
     */
    public List<Group> getGroups() throws ExecutionException, InterruptedException {
        return mDirectoryClient
                .getgroups()
                .orderBy(
                        mSnippetApp.
                                getString(
                                        R.string.displayName).
                                toString())
                .read()
                .get();
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


