/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.helpers;

import android.os.AsyncTask;

import com.microsoft.office365.snippetapp.Interfaces.IOperationCompleteListener;

public class AsyncUseCaseWrapper extends AsyncTask<BaseUserStory, IOperationCompleteListener.OperationResult, Void> {

    private final IOperationCompleteListener mOperationCompletedListener;


    public AsyncUseCaseWrapper(IOperationCompleteListener operationCompleteListener) {
        mOperationCompletedListener = operationCompleteListener;
    }

    @Override
    protected Void doInBackground(BaseUserStory... params) {
        String result;
        for (BaseUserStory useCase : params) {
            useCase.onPreExecute();
            result = useCase.execute();
            useCase.onPostExecute();
            publishProgress(new IOperationCompleteListener.OperationResult("O365 operation",
                    result));
        }
        return null;
    }

    @Override
    protected void onProgressUpdate(IOperationCompleteListener.OperationResult... values) {
        if (null != mOperationCompletedListener) {
            for (IOperationCompleteListener.OperationResult result : values) {
                mOperationCompletedListener.onOperationComplete(result);
            }
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
