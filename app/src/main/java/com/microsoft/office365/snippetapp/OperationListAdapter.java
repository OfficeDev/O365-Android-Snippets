/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

package com.microsoft.office365.snippetapp;

import android.app.Activity;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.TextView;

import com.microsoft.office365.snippetapp.helpers.BaseUserStory;

import java.util.List;

class OperationListAdapter extends BaseAdapter {

    protected LayoutInflater mLayoutInflater;
    protected List<BaseUserStory> mBaseStoryList;
    protected TextView mOperationName;
    protected View mOperationProgressBar;

    OperationListAdapter(Activity activityContext, List<BaseUserStory> stories) {
        if (null == activityContext) {
            throw new IllegalArgumentException("Context cannot be null");
        }

        if (null == stories) {
            throw new IllegalArgumentException("Use cases cannot be null");
        }
        mLayoutInflater = activityContext.getLayoutInflater();
        mBaseStoryList = stories;
    }

    @Override
    public int getCount() {
        return mBaseStoryList.size();
    }

    @Override
    public BaseUserStory getItem(int position) {
        return mBaseStoryList.get(position);
    }

    @Override
    public long getItemId(int position) {
        return 0;
    }

    @Override
    public View getView(int position, View convertView, ViewGroup parent) {
        CharSequence opDescription = getStoryDescription(position);
        int progressVisibility = isStoryExecuting(position) ? View.VISIBLE : View.INVISIBLE;

        if (!getStoryGroupingFlag(position)){
            if (null == convertView) {
                convertView = mLayoutInflater.inflate(R.layout.list_item_task, parent, false);
            }
        }
        else {
            if (null == convertView) {
                convertView = mLayoutInflater.inflate(R.layout.list_item_grouper, parent, false);
            }

        }

        if (!getStoryGroupingFlag(position)) {
            mOperationName = (TextView) convertView.findViewById(R.id.use_case_name);
            mOperationProgressBar = convertView.findViewById(R.id.use_case_progress);
            mOperationProgressBar.setVisibility(progressVisibility);
        }
        else{
            mOperationName = (TextView) convertView.findViewById(R.id.Group_name);
        }

        mOperationName.setText(opDescription);
        return convertView;
    }

    /**
     * Get the description of a story
     *
     * @param position the story to examine
     * @return the name to display
     */
    protected CharSequence getStoryDescription(int position) {
        return getItem(position).getDescription();
    }

    /**
     * Gets the story grouping flag of a story
     *
     * @param position the story to examine
     * @return the grouping flag
     */
    protected boolean getStoryGroupingFlag(int position){
        return getItem(position).getGroupingFlag();
    }
    /**
     * Check if a story is executing
     *
     * @param position the position of the use case to examine
     * @return true if the use case is executing, otherwise false
     */
    protected boolean isStoryExecuting(int position) {
        return getItem(position).isExecuting();
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
