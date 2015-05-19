/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp;

import android.app.Activity;
import android.app.ListFragment;
import android.os.Bundle;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.View;
import android.widget.BaseAdapter;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Toast;

import com.microsoft.office365.snippetapp.Interfaces.IOperationCompleteListener;
import com.microsoft.office365.snippetapp.Interfaces.O365Operations;
import com.microsoft.office365.snippetapp.Interfaces.OnUseCaseStatusChangedListener;
import com.microsoft.office365.snippetapp.helpers.AsyncUseCaseWrapper;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryList;

public class OperationListFragment extends ListFragment implements IOperationCompleteListener, OnUseCaseStatusChangedListener {

    private static final String DISCONNECTED_FROM_OFFICE_365 = "You are disconnected from Office 365";
    private static final String ON_ATTACH_EXCEPTION_MSG = "Activity must implement fragment's callbacks.";
    private StoryList mCommands;
    private O365Operations mO365Operations;
    private BaseAdapter mOperationAdapter;

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle presses on the action bar items
        switch (item.getItemId()) {
            case R.id.menu_signin:
                mO365Operations.connectToO365();
                break;
            case R.id.menu_signout:
                mO365Operations.disconnectFromO365();
                break;
            case R.id.menu_runall:
                runAllUseCases();
                break;
            case R.id.menu_clearResults:
                mO365Operations.clearResults();
                break;
            default:
                break;

        }
        return super.onOptionsItemSelected(item);
    }

    @Override
    public void onOperationComplete(final OperationResult opResult) {
        if (isAdded() && null != mO365Operations) {
            mO365Operations.getResultView().post(
                    new Runnable() {
                        @Override
                        public void run() {
                            TextView textView = (TextView) mO365Operations.getResultView();
                            textView.setText(opResult.getOperationResult() + " " + textView.getText());
                        }
                    }
            );
        }
    }

    @Override
    public void onUseCaseStatusChanged() {
        if (isAdded()) {
            getListView().post(
                    new Runnable() {
                        @Override
                        public void run() {
                            getLazyAdapter().notifyDataSetChanged();
                        }
                    }
            );
        }
    }

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setHasOptionsMenu(true);
        mCommands = new StoryList(getActivity().getApplicationContext());
    }

    @Override
    public void onViewCreated(View view, Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        setListAdapter(getLazyAdapter());
    }

    private BaseAdapter getLazyAdapter() {
        if (null == mOperationAdapter) {
            mOperationAdapter = new OperationListAdapter(getActivity(), mCommands.ITEMS);
        }
        return mOperationAdapter;
    }

    @Override
    public void onAttach(Activity activity) {
        super.onAttach(activity);

        // Activities containing this fragment must implement its callbacks.
        if (!(activity instanceof O365Operations)) {
            throw new IllegalStateException(ON_ATTACH_EXCEPTION_MSG);
        }

        mO365Operations = (O365Operations) activity;
    }

    @Override
    public void onListItemClick(ListView listView, View view, int position, long id) {
        BaseUserStory story = (BaseUserStory) listView.getAdapter().getItem(position);
        //Check that story is not a group separator for story groups
        if (!story.getGroupingFlag()) {
            //Check that app is connected to Office 365
            if (mO365Operations.getO365MailClient() == null
                    || mO365Operations.getMailServiceResourceId() == null) {
                Toast.makeText(
                        getActivity(),
                        DISCONNECTED_FROM_OFFICE_365,
                        Toast.LENGTH_LONG
                ).show();
            } else {
                //Run the story
                AsyncUseCaseWrapper asyncUseCaseWrapper = new AsyncUseCaseWrapper(this);
                asyncUseCaseWrapper.execute(story);
            }

        }
    }

    public void runAllUseCases() {
        if (mO365Operations.getO365MailClient() != null
                && mO365Operations.getO365MyFilesClient() != null) {
            AsyncUseCaseWrapper asyncUseCaseWrapper = new AsyncUseCaseWrapper(this);

            //set size with variable to avoid extra construction call with call to toArray
            int size=0;
            asyncUseCaseWrapper.execute(mCommands.ITEMS.toArray(new BaseUserStory[size]));
        } else {
            Toast.makeText(
                    getActivity(),
                    DISCONNECTED_FROM_OFFICE_365,
                    Toast.LENGTH_LONG
            ).show();
        }
    }

    public void onServicesReady() {
        for (BaseUserStory story : mCommands.ITEMS) {
            story.setO365MailClient(mO365Operations.getO365MailClient());
            story.setO365MailResourceId(mO365Operations.getMailServiceResourceId());
            story.setO365MyFilesClient(mO365Operations.getO365MyFilesClient());
            story.setFilesFoldersResourceId(mO365Operations.getMyFilesServiceResourceId());
            story.setUIResultView(mO365Operations.getResultView());
            story.setUseCaseStatusChangedListener(this);
        }
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
