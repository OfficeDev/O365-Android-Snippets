/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.Snippets;

import android.util.Patterns;

import com.microsoft.outlookservices.Attendee;
import com.microsoft.outlookservices.BodyType;
import com.microsoft.outlookservices.EmailAddress;
import com.microsoft.outlookservices.Event;
import com.microsoft.outlookservices.ItemBody;
import com.microsoft.outlookservices.ResponseStatus;
import com.microsoft.outlookservices.ResponseType;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.regex.Matcher;

public class CalendarSnippets {

    OutlookClient mCalendarClient;

    public CalendarSnippets(OutlookClient mailClient) {
        mCalendarClient = mailClient;
    }

    //Return a list of calendar events off the default calendar ordered by the Start field.
    public List<Event> getO365Events() throws ExecutionException, InterruptedException {

        int EVENT_RANGE_START = 1;
        int EVENT_RANGE_END = 1;
        int PAGE_SIZE = 10;
        String SORT_COLUMN = "Start";
        java.util.Calendar dateStart = java.util.Calendar.getInstance();

        //Set the date range of the calendar view to retrieve
        dateStart.add(Calendar.WEEK_OF_MONTH, -EVENT_RANGE_START);
        java.util.Calendar dateEnd = java.util.Calendar.getInstance();
        dateEnd.add(Calendar.WEEK_OF_MONTH, EVENT_RANGE_END);

        return mCalendarClient
                .getMe()
                .getCalendarView()
                .addParameter("startdatetime", dateStart)
                .addParameter("enddatetime", dateEnd)
                .top(PAGE_SIZE)
                .orderBy(SORT_COLUMN)
                .read()
                .get();

    }

    public void deleteCalendarEvent(String eventId)
            throws ExecutionException
            , InterruptedException {
        mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId).delete().get();

    }

    public String createCalendarEvent(
            String subject
            , String itemBodyHtml
            , java.util.Calendar startDate
            , Calendar endDate
            , List<String> attendees)
            throws ExecutionException
            , InterruptedException {
        Event newEvent = new Event();
        newEvent.setSubject(subject);
        ItemBody itemBody = new ItemBody();
        itemBody.setContent(itemBodyHtml);
        itemBody.setContentType(BodyType.HTML);
        newEvent.setBody(itemBody);
        newEvent.setStart(startDate);
        newEvent.setIsAllDay(false);
        newEvent.setEnd(endDate);
        Matcher matcher;
        List<Attendee> attendeeList = new ArrayList<>();
        for (String s : attendees) {
            // Add mail to address if mailToString is an email address
            matcher = Patterns.EMAIL_ADDRESS.matcher(s);
            if (matcher.matches()) {
                EmailAddress emailAddress = new EmailAddress();
                emailAddress.setAddress(s);
                Attendee attendee = new Attendee();
                attendee.setEmailAddress(emailAddress);
                attendeeList.add(attendee);
            }
        }
        newEvent.setAttendees(attendeeList);


        return mCalendarClient
                .getMe()
                .getEvents()
                .add(newEvent).get().getId();
    }

    public Event updateCalendarEvent(
            String eventId
            , String subject
            , boolean allDay
            , String itemBodyHtml
            , java.util.Calendar startDate
            , Calendar endDate
            , List<String> attendees) throws ExecutionException, InterruptedException {
        Event calendarEvent = getCalendarEvent(eventId);
        calendarEvent.setSubject(subject);
        if (itemBodyHtml != null && itemBodyHtml.length() > 0) {
            ItemBody itemBody = new ItemBody();
            itemBody.setContent(itemBodyHtml);
            itemBody.setContentType(BodyType.HTML);
            calendarEvent.setBody(itemBody);
        }
        if (startDate != null) {
            calendarEvent.setStart(startDate);
        }
        if (endDate != null) {
            calendarEvent.setEnd(endDate);
        }

        calendarEvent.setIsAllDay(allDay);

        if (attendees != null && attendees.size() > 0) {
            //clear attendee list and set with new list
            calendarEvent.setAttendees(null);
            Matcher matcher;
            List<Attendee> attendeeList = new ArrayList<>();
            for (String s : attendees) {
                // Add mail to address if mailToString is an email address
                matcher = Patterns.EMAIL_ADDRESS.matcher(s);
                if (matcher.matches()) {
                    EmailAddress emailAddress = new EmailAddress();
                    emailAddress.setAddress(s);
                    Attendee attendee = new Attendee();
                    attendee.setEmailAddress(emailAddress);
                    attendeeList.add(attendee);
                }
            }
            calendarEvent.setAttendees(attendeeList);
        }

        return mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId)
                .update(calendarEvent)
                .get();

    }

    public String getEventAttendeeStatus(String eventId, String myEmailAddress) {
        for (Attendee attendee : getCalendarEvent(eventId).getAttendees()) {
            String attendeeEmail = attendee.getEmailAddress().getAddress();
            if (attendeeEmail.equalsIgnoreCase(myEmailAddress)) {
                return attendee.getStatus().getResponse().toString();
            }
        }
        return null;
    }

    public Event acceptCalendarEventInvite(String eventId, String myEmailAddress) throws ExecutionException, InterruptedException {
        Event calendarEvent = getCalendarEvent(eventId);

        for (Attendee a : calendarEvent.getAttendees()) {
            String thisEmailAddress = a.getEmailAddress().getAddress();
            if (thisEmailAddress.toLowerCase().equals(myEmailAddress.toLowerCase())) {
                ResponseStatus inviteResponse = new ResponseStatus();
                inviteResponse.setResponse(ResponseType.Accepted);
                a.setStatus(inviteResponse);
                break;
            }

        }
        return mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId)
                .update(calendarEvent)
                .get();
    }

    public Event declineCalendarEventInvite(String eventId, String myEmailAddress) throws ExecutionException, InterruptedException {
        Event calendarEvent = getCalendarEvent(eventId);

        if (calendarEvent.getAttendees().size() > 0) {
            for (Attendee a : calendarEvent.getAttendees()) {
                String thisEmailAddress = a.getEmailAddress().getAddress();
                if (thisEmailAddress.toLowerCase().equals(myEmailAddress.toLowerCase())) {
                    ResponseStatus inviteResponse = new ResponseStatus();
                    inviteResponse.setResponse(ResponseType.Declined);
                    a.setStatus(inviteResponse);
                    break;
                }
            }
        }

        return mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId)
                .update(calendarEvent)
                .get();
    }

    public String getCalendarEventId(String eventId) throws ExecutionException, InterruptedException {
        return mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId)
                .read()
                .get()
                .getId();
    }

    public Event getCalendarEvent(String eventId) {
        try {
            return mCalendarClient
                    .getMe()
                    .getEvents()
                    .getById(eventId)
                    .read()
                    .get();
        } catch (Exception e) {
            return null;
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
