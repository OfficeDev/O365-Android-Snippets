/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import com.microsoft.outlookservices.Contact;
import com.microsoft.outlookservices.EmailAddress;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class ContactsSnippets {

    private final OutlookClient mOutlookClient;

    public ContactsSnippets(OutlookClient mailClient) {
        mOutlookClient = mailClient;
    }

    /**
     * Return a list of contacts and ordered by the
     * contact's surname field.
     *
     * @return List. A list of the com.microsoft.outlookservices.Contact objects
     * @version 1.0
     */
    public List<Contact> getContacts(int pageSize) throws ExecutionException, InterruptedException {
        return mOutlookClient
                .getMe()
                .getContacts()
                .top(pageSize)
                .skip(1)
                .orderBy("SurName")
                .read().get();
    }

    /**
     * Creates a new contact
     *
     * @param emailAddressString  The email address of the contact to be added
     * @param businessPhoneString The business telephone number of the new contact
     * @param firstNameString     The first name of the new contact
     * @param homePhoneString     The home telephone number of the new contact
     * @param lastNameString      The surname of the new contact
     * @return String. The id of the new contact
     * @version 1.0
     */
    public String createContact(
            String emailAddressString,
            String businessPhoneString,
            String homePhoneString,
            String firstNameString,
            String lastNameString) throws ExecutionException, InterruptedException {
        Contact contact = new Contact();

        EmailAddress emailAddress = new EmailAddress();
        emailAddress.setAddress(emailAddressString);
        List<EmailAddress> emailAddressList = new ArrayList<>();
        emailAddressList.add(emailAddress);
        contact.setEmailAddresses(emailAddressList);

        List<String> businessPhones = new ArrayList<>();
        businessPhones.add(businessPhoneString);
        List<String> homePhones = new ArrayList<>();
        homePhones.add(homePhoneString);
        contact.setBusinessPhones(businessPhones);
        contact.setHomePhones(homePhones);
        contact.setGivenName(firstNameString);
        contact.setSurname(lastNameString);

        return mOutlookClient
                .getMe()
                .getContacts()
                .add(contact).get().getId();
    }

    /**
     * Gets a contact by the contact Id
     *
     * @return Contact. The contact corresponding to the id
     * @version 1.0
     */
    public Contact getAContact(String id) throws ExecutionException, InterruptedException {
        return mOutlookClient
                .getMe()
                .getContacts()
                .getById(id).read().get();
    }


    /**
     * Updates the first and surname of a contact
     *
     * @param contactId       The id of the contact to be updated
     * @param firstNameString The updated first name of the new contact
     * @param lastNameString  The updated surname of the new contact
     * @version 1.0
     */
    public void updateContact(String contactId, String firstNameString, String lastNameString)
            throws ExecutionException, InterruptedException {
        //Get the contact to update
        Contact updateContact = mOutlookClient
                .getMe()
                .getContacts()
                .getById(contactId)
                .read()
                .get();

        //Update the contact first and last name
        updateContact.setSurname(lastNameString);
        updateContact.setGivenName(firstNameString);

        //push the updated contact to the server
        mOutlookClient
                .getMe()
                .getContacts()
                .getById(contactId)
                .update(updateContact).get();


    }

    /**
     * Deletes a contact
     *
     * @param id The id of the contact to be deleted
     * @throws ExecutionException
     * @throws InterruptedException
     */
    public void deleteContact(String id) throws ExecutionException, InterruptedException {
        mOutlookClient
                .getMe()
                .getContacts()
                .getById(id)
                .addHeader("If-Match", "*")
                .delete()
                .get();

    }

    /**
     * Runs a filtered query to find all contacts with the given surname. This snippet can be
     * modified to run any filtered query. For a complete list of contact properties that
     * can be filtered, see https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesEvent
     *
     * @param surname
     * @return A list of contacts matching the given surname
     * @throws ExecutionException
     * @throws InterruptedException
     */
    public List<Contact> getContactsWithSurname(String surname) throws ExecutionException, InterruptedException {
        StringBuilder query = new StringBuilder();
        query.append("surname eq '")
                .append(surname)
                .append("'");
        return mOutlookClient
                .getMe()
                .getContacts()
                .filter(query.toString())
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
