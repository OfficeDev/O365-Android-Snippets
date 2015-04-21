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

    /**
     * Return a list of calendar events from the default calendar and ordered by the
     * Start field. The range of events to be returned includes events from
     * 1 week prior to current date through 1 week into the future. 10 events
     * are returned per page
     *
     * @return List. A list of the com.microsoft.outlookservices.Event objects
     * @version 1.0
     */
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

    /**
     * Removes an event specified by the id
     * are returned per page
     *
     * @param eventId The id of the event to be removed
     * @version 1.0
     */
    public void deleteCalendarEvent(String eventId)
            throws ExecutionException
            , InterruptedException {
        mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId).delete().get();

    }

    /**
     * Creates an event
     *
     * @param subject      The subject of the event
     * @param itemBodyHtml The body of the event as HTML
     * @param startDate    The start date of the event
     * @param endDate      The end date of the event
     * @param attendees    A list of attendee email addresses
     * @return String The id of the created event
     * @version 1.0
     */
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

    /**
     * Updates the subject, body, start date, end date, or attendees of an event
     *
     * @param subject      The subject of the event
     * @param allDay       true if event spans a work day
     * @param itemBodyHtml The body of the event as HTML
     * @param startDate    The start date of the event
     * @param endDate      The end date of the event
     * @param attendees    A list of attendee email addresses
     * @return String The id of the created event
     * @version 1.0
     */
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

    /**
     * Create a calendar event based off a full Event object passed to this method.
     *
     * @return the same Event object updated with the ID from the server.
     */
    public Event createCalendarEvent(Event eventToCreate) throws ExecutionException, InterruptedException {
        return mCalendarClient
                .getMe()
                .getEvents()
                .add(eventToCreate).get();
    }

    /**
     * Gets the invitation status of a given attendee for a given event
     *
     * @param eventId        The id of the event to be removed
     * @param myEmailAddress The email address of the attendee whose status is of interest
     * @version 1.0
     */
    public String getEventAttendeeStatus(String eventId, String myEmailAddress) throws ExecutionException, InterruptedException {
        for (Attendee attendee : getCalendarEvent(eventId).getAttendees()) {
            String attendeeEmail = attendee.getEmailAddress().getAddress();
            if (attendeeEmail.equalsIgnoreCase(myEmailAddress)) {
                return attendee.getStatus().getResponse().toString();
            }
        }
        return null;
    }

    /**
     * Accepts an event invitation on behalf of the specified attendee
     *
     * @param eventId        The id of the event to be removed
     * @param myEmailAddress The email address of the attendee whose status is of interest
     * @version 1.0
     */
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

    /**
     * Declines an event invitation on behalf of the specified attendee
     *
     * @param eventId        The id of the event to be responded to
     * @param myEmailAddress The email address of the attendee whose status is of interest
     * @version 1.0
     */
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

    /**
     * Gets the event id of the event specified by an event Id
     *
     * @param eventId The id of the event to be responded to
     * @version 1.0
     */
    public String getCalendarEventId(String eventId) throws ExecutionException, InterruptedException {
        return mCalendarClient
                .getMe()
                .getEvents()
                .getById(eventId)
                .read()
                .get()
                .getId();
    }

    /**
     * Gets the event object of the event specified by an event Id
     *
     * @param eventId The id of the event to be responded to
     * @return Event The event of interest
     * @version 1.0
     */
    public Event getCalendarEvent(String eventId) throws ExecutionException, InterruptedException {
            return mCalendarClient
                    .getMe()
                    .getEvents()
                    .getById(eventId)
                    .read()
                    .get();
    }

    /**
     * Runs a filtered query to find all events that are high importance. This snippet can be
     * modified to run any filtered query. For a complete list of Events properties that
     * can be filtered, see https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesEvent
     *
     * @return A list of events
     * @version 1.0
     */
    public List<Event> getImportantEvents() throws ExecutionException, InterruptedException {
        return mCalendarClient
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
