/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.helpers;

import com.microsoft.directoryservices.odata.DirectoryClient;
import com.microsoft.outlookservices.odata.OutlookClient;
import com.microsoft.services.odata.impl.ADALDependencyResolver;
import com.microsoft.services.odata.interfaces.DependencyResolver;

public class O365ServicesManager {
    static DirectoryClient mDirectoryClient=null;
    static String mTenantId=null;

    public static void initialize(String tenantId){
        mTenantId=tenantId;
    }

    public static DirectoryClient getDirectoryClient(){
        if (mDirectoryClient==null&&mTenantId!=null){
            DependencyResolver dependencyResolver = AuthenticationController.getInstance().getDependencyResolver();
            StringBuilder endpoint = new StringBuilder();
            endpoint.append(Constants.DIRECTORY_RESOURCE_URL)
                    .append(mTenantId)
                    .append("?")
                    .append(Constants.DIRECTORY_API_VERSION);
            mDirectoryClient=new DirectoryClient(endpoint.toString(),dependencyResolver);
        }
        return mDirectoryClient;
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
