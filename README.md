# SPCalendarPro
An ultra lightweight, dependency-free JavaScript library to easily manage SharePoint calendar events.


## Purpose
This library is intended to allow users to make dealing with SharePoint calendars much more simple. 

The painful process of obtaining recurring events, matching user provided datetimes with events, and determining time conflicts is tortorous and requires multiple dependencies and lots of code. SP Calendar Pro simplifies this process enormously.


## Features
1) Easily collects calendar event items (recurring events, single events, or a combination of the two).
2) Provides simple way to collect requested datetime values the user. End user can either pass the datetime variables into a method, or have the requested datetime values automatically be extracted from a SharePoint form.
3) Provides several methods to easily facilitate various datetime comparisons (match specified datetimes, determine time conflicts, same dates, etc).
4) Requires zero dependencies!



## Example

This example below will:

a) Collect all events (single and recurring) from the 'Appointments' calendar list. 

b) Convert user provided datetime information from a SharePoint form into proper date objects for use. The "0,1" parameters specify which datetime field elements on a form are to be converted.

c) Determine if user provided datetime values pose a time conflict with any events from the Appointments list.

d) Filters the returned events to match only the ones with a Status of 'Confirmed'.

    var timeConflicts = spcalpro.getEvents('Appointments', false, '').getDateTimesFromForm(0,1).isTimeConflict().where('Status = Confirmed');

This returns an object that contains only the events that have time conflicts with the requested datetimes:

    for (var i = 0; i < timeConflicts.events.length; i++) {
        console.log( timeConflicts.events[i] );
    }

### This library is still in its pre-beta phase, so you can expect numerous updates, revisions, bug fixes, and improvements over the coming weeks.