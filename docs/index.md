# SPCalendarPro
An ultra lightweight, dependency-free JavaScript library to easily manage SharePoint calendar events.

* 15 KB unminified, with comments.
* 7 KB minified.


## Purpose
This library is intended to allow users to simplify dealing with SharePoint calendars. 

The painful process of obtaining recurring events, matching user provided datetimes with events, and determining time conflicts is tortorous and requires multiple dependencies and lots of code. SP Calendar Pro simplifies this process enormously.


## Features
1) Easily collects calendar event items (recurring events, single events, or a combination of the two).
2) Provides a simple way to collect requested datetime values the user. You can either pass the datetime variables into a method, or use a method to convert datetimes from a SharePoint form into proper datetime variables.
3) Provides several methods to easily facilitate various datetime comparisons. Match specified datetimes: matchDateTimes(), determine time conflicts: isTimeConflict(), same dates: isSameDate(), etc).\
4) A basic 'where' operator to provide simple filtering based on field values.
5) Return only datetimes that occur after today: getEventsAfterToday().
6) Returns list items along with calendar events.
7) Option to disable calendar drag and drop: spcalpro.disableDragAndDrop()
8) Requires zero dependencies! Everything is pure vanilla JS.
9) Compatible down to SharePoint 2010.



## Example

This example below will:

a) Asynchronously collect all events (single and recurring) from the 'Appointments' calendar list. 

b) Return only the events after today.

c) Convert user provided datetime information from a SharePoint form into proper date objects for use. The "0,1" parameters specify which datetime field elements on a form are to be converted.

d) Determine if user provided datetime values pose a time conflict with any events from the Appointments list.

e) Filters the returned events to match only the ones with a Status of 'Confirmed'.

f) The above three steps are executed in the callback method, which is getEvents().

    var data = {
        listName: 'Appointments',
        async: true,
        type: '',
        getEventsAfterDate: new Date(),
        callback: getEvents,
    }

    spcalpro.getCalendarEvents(data);

    function getEvents(obj) {
        var events = obj.getDateTimesFromForm(0,1).isTimeConflict().where('Status = Confirmed');
        console.table(events.listData);
    }
    

This will display all events from the Appointments calendar after today, which pose a time conflict with the user specified date times from a form, and where the events' Status is Confirmed.


### This library is still in its beta phase, so you can expect numerous updates, revisions, bug fixes, and improvements over the coming weeks.