/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import com.microsoft.outlookservices.Event;
import com.microsoft.outlookservices.Message;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class ODataSystemQuerySnippets {

    /**
     * Demonstrates how to specify the $select OData system query option.
     * Gets a list of the 10 email messages selecting only the from, subject,
     * and isRead fields. Use the .select() method to specify a $select query
     * as shown in this snippet.
     * <p/>
     * For a complete list of types and properties that can be selected
     * see https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#OdataQueryParams
     *
     * @return List of type com.microsoft.outlookservices.Message
     * @version 1.0
     */
    public List<Message> getMailMessagesUsing$select(OutlookClient outlookClient) throws ExecutionException, InterruptedException {
        List<Message> messages = outlookClient
                .getMe()
                .getFolders()
                .getById("Inbox")
                .getMessages()
                .select("from,subject,isRead")
                .top(10)
                .read().get();
        return messages;
    }

    /**
     * Demonstrates how to specify the $filter OData system query option. Runs a filtered
     * query to find all events that are high importance. Use the .filter() method
     * to specify a $filter query as shown in this snippet.
     * <p/>
     * For a complete list of types and properties that can be filtered
     * see https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#OdataQueryParams
     *
     * @param outlookClient - A client providing access to Office 365 Outlook APIs
     * @return A list of important events
     * @throws ExecutionException
     * @throws InterruptedException
     */
    public List<Event> getImportantEventsUsing$filter(OutlookClient outlookClient) throws ExecutionException, InterruptedException {
        return outlookClient
                .getMe()
                .getEvents()
                .filter("Importance eq 'High'")
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
