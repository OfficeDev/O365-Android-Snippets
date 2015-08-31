/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
* Portions of this class are adapted from the AuthenticationController.java file from Microsoft Open Technologies, Inc.
* located at https://github.com/OfficeDev/Office-365-SDK-for-Android/blob/master/samples/outlook/app/src/main/java/com/microsoft/services/controllers/AuthenticationController.java
*/
package com.microsoft.office365.snippetapp.helpers;

import android.app.Activity;
import android.util.Log;

import com.google.common.util.concurrent.SettableFuture;
import com.microsoft.aad.adal.AuthenticationCallback;
import com.microsoft.aad.adal.AuthenticationContext;
import com.microsoft.aad.adal.AuthenticationResult;
import com.microsoft.aad.adal.AuthenticationResult.AuthenticationStatus;
import com.microsoft.aad.adal.PromptBehavior;
import com.microsoft.services.orc.core.DependencyResolver;
import com.microsoft.services.orc.log.LogLevel;
import com.microsoft.services.orc.resolvers.ADALDependencyResolver;

/**
 * Handles setup of ADAL Dependency Resolver for use in API clients.
 */

public class AuthenticationController {
    private static String TAG = "AuthenticationController";
    private static AuthenticationController INSTANCE;
    private AuthenticationContext authContext;
    private ADALDependencyResolver dependencyResolver;
    private Activity contextActivity;
    private String resourceId;
    private String mAccessToken;

    private AuthenticationController() {
        resourceId = com.microsoft.office365.snippetapp.helpers.Constants.DISCOVERY_RESOURCE_ID;
    }

    private AuthenticationController(final Activity contextActivity) {
        this();
        this.contextActivity = contextActivity;
    }

    public static synchronized AuthenticationController getInstance() {
        if (INSTANCE == null) {
            INSTANCE = new AuthenticationController();
        }
        return INSTANCE;
    }

    public static synchronized void resetInstance() {
        INSTANCE = null;
    }

    /**
     * Set the context activity before initializing to the currently active activity.
     *
     * @param contextActivity Currently active activity which can be utilized for interactive
     *                        prompt.
     */
    public void setContextActivity(final Activity contextActivity) {
        this.contextActivity = contextActivity;
    }

    /**
     * Change from the default Resource ID set in ServiceConstants to a different
     * resource ID.
     * This can be called at anytime without requiring another interactive prompt.
     *
     * @param resourceId URL of resource ID to be accessed on behalf of user.
     */
    public void setResourceId(final String resourceId) {
        this.resourceId = resourceId;
        this.dependencyResolver.setResourceId(resourceId);
    }

    /**
     * Returns the access token for the connected user
     * @return String
     */
    public String getAccessToken(){
        return mAccessToken;
    }
    /**
     * Turn on logging.
     *
     * @param level LogLevel to set.
     */
    public void enableLogging(LogLevel level) {
        this.dependencyResolver.getLogger().setEnabled(true);
        this.dependencyResolver.getLogger().setLogLevel(level);
    }

    /**
     * Turn off logging.
     */
    public void disableLogging() {
        this.dependencyResolver.getLogger().setEnabled(false);
    }

    /**
     * Description: Calls AuthenticationContext.acquireToken(...) once to initialize with
     * user's credentials and avoid interactive prompt on later calls.
     * If all tokens expire, app must call initialize() again to prompt user interactively and
     * set up authentication context.
     *
     * @return A signal to wait on before continuing execution.
     */
    public SettableFuture<AuthenticationResult> initialize() {

        final SettableFuture<AuthenticationResult> result = SettableFuture.create();

        if (verifyAuthenticationContext()) {
            AuthenticationContext authContext = getAuthenticationContext();
            if (authContext != null)
                authContext.acquireToken(
                        this.contextActivity,
                        this.resourceId,
                        Constants.CLIENT_ID,
                        Constants.REDIRECT_URI,
                        PromptBehavior.Auto,
                        new AuthenticationCallback<AuthenticationResult>() {

                            @Override
                            public void onSuccess(final AuthenticationResult authenticationResult) {

                                if (authenticationResult != null && authenticationResult.getStatus() == AuthenticationStatus.Succeeded) {
                                    dependencyResolver = new ADALDependencyResolver(
                                            getAuthenticationContext(),
                                            resourceId,
                                            Constants.CLIENT_ID);
                                    O365ServicesManager.initialize(authenticationResult.getTenantId());
                                    mAccessToken = authenticationResult.getAccessToken();
                                    result.set(authenticationResult);
                                }
                            }

                            @Override
                            public void onError(Exception t) {
                                result.setException(t);
                            }
                        }
                );
            else
                result.setException(new Throwable("Auth context verification failed. Did you set a context activity?"));
        } else {
            result.setException(new Throwable("Auth context verification failed. Did you set a context activity?"));
        }
        return result;
    }

    /**
     * Gets AuthenticationContext for AAD.
     *
     * @return authenticationContext, if successful.
     */
    public AuthenticationContext getAuthenticationContext() {
        if (authContext == null) {
            try {
                authContext = new AuthenticationContext(
                        this.contextActivity
                        , Constants.AUTHORITY_URL
                        , false);

            } catch (Throwable t) {
                Log.e(TAG, t.toString());
            }
        }
        return authContext;
    }

    public DependencyResolver getDependencyResolver() {
        return getInstance().dependencyResolver;
    }

    private boolean verifyAuthenticationContext() {
        if (this.contextActivity == null) {
            Log.e(TAG, "Must set context activity");
            return false;
        }
        return true;
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
