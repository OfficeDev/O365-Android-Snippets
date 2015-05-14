/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp;

import android.app.Activity;
import android.app.FragmentTransaction;
import android.content.Intent;
import android.os.Bundle;
import android.os.PersistableBundle;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.View;
import android.webkit.CookieSyncManager;
import android.widget.Toast;

import com.google.common.util.concurrent.FutureCallback;
import com.google.common.util.concurrent.Futures;
import com.google.common.util.concurrent.SettableFuture;
import com.microsoft.aad.adal.AuthenticationResult;
import com.microsoft.discoveryservices.ServiceInfo;
import com.microsoft.fileservices.odata.SharePointClient;
import com.microsoft.office365.snippetapp.Interfaces.O365Operations;
import com.microsoft.office365.snippetapp.helpers.AuthUtil;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.Constants;
import com.microsoft.office365.snippetapp.helpers.DiscoveryController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.outlookservices.odata.OutlookClient;
import com.microsoft.services.odata.impl.ADALDependencyResolver;

import java.net.URI;
import java.util.UUID;

import static com.microsoft.office365.snippetapp.R.id.operation_detail_container;


public class OperationListActivity extends Activity
        implements O365Operations {

    public static final String DISCONNECTED_FROM_OFFICE = "You are disconnected from Office 365";
    public static final int SIGNIN_MENU_ITEM = 1;
    public static final int SIGNOUT_MENU_ITEM = 2;
    private static final String TAG = "OperationListActivity";
    public OutlookClient mOutlookClient;
    public SharePointClient mMyFilesClient;
    private String mMailServiceResourceId;
    private String mMailServiceEndpointUri;
    private String mMyFilesServiceEndpointUri;
    private String mMyFilesServiceResourceId;
    private MenuItem mSignIn;
    private MenuItem mSignOut;
    private OperationDetailFragment mOperationDetailFragment;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_operation_list);
        if (findViewById(operation_detail_container) != null) {

            mOperationDetailFragment = new OperationDetailFragment();
            FragmentTransaction fragmentTransaction = getFragmentManager()
                    .beginTransaction()
                    .add(R.id.operation_detail_container, mOperationDetailFragment);

            fragmentTransaction.commit();
        }
        AuthUtil.configureAuthSettings(this);
    }

    @Override
    protected void onPause() {
        super.onPause();
    }

    @Override
    protected void onResume() {
        super.onResume();
    }

    @Override
    public void onSaveInstanceState(Bundle outState, PersistableBundle outPersistentState) {
        super.onSaveInstanceState(outState, outPersistentState);
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater menuInflater = getMenuInflater();
        menuInflater.inflate(R.menu.main, menu);
        mSignIn = menu.getItem(SIGNIN_MENU_ITEM);
        mSignOut = menu.getItem(SIGNOUT_MENU_ITEM);
        return super.onCreateOptionsMenu(menu);
    }

    @Override
    protected void onRestoreInstanceState(Bundle savedInstanceState) {
        super.onRestoreInstanceState(savedInstanceState);
    }

    private void enableListClick() {
        this.runOnUiThread(
                new Runnable() {
                    @Override
                    public void run() {
                        OperationListFragment operationListFragment = (OperationListFragment) getFragmentManager()
                                .findFragmentById(R.id.operation_list);
                        operationListFragment.onServicesReady();
                    }
                }
        );
    }

    @Override
    public void connectToO365() {
        //check that client id and redirect have been set correctly
        try {
            UUID.fromString(Constants.CLIENT_ID);
            URI.create(Constants.REDIRECT_URI);
        } catch (IllegalArgumentException e) {
            Toast.makeText(
                    this
                    , getString(R.string.warning_clientid_redirecturi_incorrect)
                    , Toast.LENGTH_LONG
            ).show();
            return;
        }


        AuthenticationController.getInstance().setContextActivity(this);
        SettableFuture<AuthenticationResult> future = AuthenticationController
                .getInstance()
                .initialize();

        Futures.addCallback(
                future, new FutureCallback<AuthenticationResult>() {
                    /**
                     * If the connection is successful, the activity extracts the username and
                     * displayableId values from the authentication result object and sends them
                     * to the SendMail activity.
                     * @param result The authentication result object that contains information about
                     *               the user and the tokens.
                     */
                    @Override
                    public void onSuccess(AuthenticationResult result) {
                        getMailServiceResource();
                        getFileFolderServiceResource();

                        //obtain email address for logged in user to use for email snippets
                        GlobalValues.USER_EMAIL = result.getUserInfo().getDisplayableId();
                        GlobalValues.USER_NAME = result.getUserInfo().getGivenName() + "  " + result.getUserInfo().getFamilyName();
                        mSignOut.setTitle("Disconnect " +  GlobalValues.USER_NAME);

                    }

                    @Override
                    public void onFailure(final Throwable t) {
                        Log.e(TAG, "onCreate - " + t.getMessage());
                        showConnectErrorUI();
                    }
                }
        );
    }

    @Override
    public void disconnectFromO365() {
        if (AuthenticationController.getInstance().getAuthenticationContext() != null) {
            AuthenticationController
                    .getInstance()
                    .getAuthenticationContext()
                    .getCache()
                    .removeAll();


            //IMPORTANT: This code removes all cookies stored for all webviews on this app
            CookieSyncManager syncManager = CookieSyncManager.createInstance(this);
            if (syncManager != null) {
                android.webkit.CookieManager cookieManager = android
                        .webkit
                        .CookieManager
                        .getInstance();
                cookieManager.removeAllCookie();
                syncManager.sync();

            }
            //Disable the Connect option on action bar, enable the disconnect option
            mSignIn.setEnabled(true);
            mSignOut.setEnabled(false);

            mMyFilesClient = null;
            mOutlookClient = null;
            Toast.makeText(
                    OperationListActivity.this,
                    DISCONNECTED_FROM_OFFICE,
                    Toast.LENGTH_LONG).show();

            //Clear last connected user name from Disconnect action menu item
            mSignOut.setTitle("Disconnect");

        }

    }

    @Override
    public void clearResults() {
        mOperationDetailFragment.clearResults();
    }

    /**
     * This activity gets notified about the completion of the ADAL activity through this method.
     *
     * @param requestCode The integer request code originally supplied to startActivityForResult(),
     *                    allowing you to identify who this result came from.
     * @param resultCode  The integer result code returned by the child activity through its
     *                    setResult().
     * @param data        An Intent, which can return result data to the caller (various data
     *                    can be attached to Intent "extras").
     */
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        Log.i(TAG, "onActivityResult - AuthenticationActivity has come back with results");
        super.onActivityResult(requestCode, resultCode, data);
        AuthenticationController
                .getInstance()
                .getAuthenticationContext()
                .onActivityResult(requestCode, resultCode, data);
    }


    @Override
    public OutlookClient getO365MailClient() {
        return mOutlookClient;
    }

    @Override
    public String getMailServiceResourceId() {
        return mMailServiceResourceId;
    }

    @Override
    public SharePointClient getO365MyFilesClient() {
        return mMyFilesClient;
    }

    @Override
    public String getMyFilesServiceResourceId() {
        return mMyFilesServiceResourceId;
    }

    @Override
    public View getResultView() {
        return findViewById(R.id.operation_detail);
    }


    private void getFileFolderServiceResource() {
        final SettableFuture<ServiceInfo> serviceDiscovered;


        serviceDiscovered = DiscoveryController
                .getInstance()
                .getServiceInfo(Constants.MYFILES_CAPABILITY);

        Futures.addCallback(
                serviceDiscovered,
                new FutureCallback<ServiceInfo>() {
                    @Override
                    public void onSuccess(ServiceInfo serviceInfo) {
                        Log.i(TAG, "onConnect - My files service discovered");
                        showDiscoverSuccessUI();

                        mMyFilesServiceResourceId =
                                serviceInfo.getserviceResourceId();

                        mMyFilesServiceEndpointUri =
                                serviceInfo.getserviceEndpointUri();

                        AuthenticationController.getInstance()
                                .setResourceId(mMyFilesServiceResourceId);
                        ADALDependencyResolver dependencyResolver = (ADALDependencyResolver) AuthenticationController
                                .getInstance()
                                .getDependencyResolver();

                        mMyFilesClient = new SharePointClient(
                                mMyFilesServiceEndpointUri,
                                dependencyResolver
                        );

                        //User is connected, SharePoint client endpoint is found
                        // Enable list item selection on main activity
                        enableListClick();
                        runOnUiThread(
                                new Runnable() {
                                    @Override
                                    public void run() {
                                        mSignIn.setEnabled(false);
                                        mSignOut.setEnabled(true);
                                    }
                                }
                        );

                    }

                    @Override
                    public void onFailure(final Throwable t) {
                        Log.e(TAG, "onConnect - " + t.getMessage());
                        showDiscoverErrorUI();
                    }
                }
        );
    }

    private void getMailServiceResource() {
        final SettableFuture<ServiceInfo> serviceDiscovered;

        serviceDiscovered = DiscoveryController
                .getInstance()
                .getServiceInfo(Constants.MAIL_CAPABILITY);

        Futures.addCallback(
                serviceDiscovered,
                new FutureCallback<ServiceInfo>() {
                    @Override
                    public void onSuccess(ServiceInfo serviceInfo) {
                        Log.i(TAG, "onConnect - Mail service discovered");
                        showDiscoverSuccessUI();

                        mMailServiceResourceId =
                                serviceInfo.getserviceResourceId();

                        mMailServiceEndpointUri =
                                serviceInfo.getserviceEndpointUri();

                        lazyMailClientGetter();
                        enableListClick();
                        runOnUiThread(
                                new Runnable() {
                                    @Override
                                    public void run() {
                                        mSignIn.setEnabled(false);
                                        mSignOut.setEnabled(true);
                                    }
                                }
                        );

                    }

                    @Override
                    public void onFailure(final Throwable t) {
                        Log.e(TAG, "onConnect - " + t.getMessage());
                        showDiscoverErrorUI();
                    }
                }
        );
    }

    private void lazyMailClientGetter() {
        AuthenticationController.getInstance()
                .setResourceId(mMailServiceResourceId);

        ADALDependencyResolver dependencyResolver = (ADALDependencyResolver) AuthenticationController
                .getInstance()
                .getDependencyResolver();
        if (mOutlookClient == null) {
            mOutlookClient = new OutlookClient(
                    mMailServiceEndpointUri,
                    dependencyResolver);
        }
    }

    private void showDiscoverSuccessUI() {
        runOnUiThread(
                new Runnable() {
                    @Override
                    public void run() {
                        Toast.makeText(
                                OperationListActivity.this,
                                R.string.discover_toast_text,
                                Toast.LENGTH_SHORT
                        ).show();
                    }
                }
        );
    }

    private void showDiscoverErrorUI() {
        runOnUiThread(
                new Runnable() {
                    @Override
                    public void run() {
                        Toast.makeText(
                                OperationListActivity.this,
                                R.string.discover_toast_text_error,
                                Toast.LENGTH_LONG
                        ).show();
                    }
                }
        );
    }

    public void showEncryptionKeyErrorUI() {
        Toast.makeText(
                OperationListActivity.this,
                R.string.encryption_key_text_error,
                Toast.LENGTH_LONG
        ).show();
    }

    private void showConnectErrorUI() {
        Toast.makeText(
                OperationListActivity.this,
                R.string.connect_toast_text_error,
                Toast.LENGTH_LONG
        ).show();
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
