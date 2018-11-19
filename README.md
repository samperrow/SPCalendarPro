# SPCalendarPro
An ultra lightweight, dependency-free JavaScript library to easily manage SharePoint calendar events.

* 16 KB unminified, with comments.
* 8 KB minified.


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
9) Compatible for all SharePoint versions- 2010, 2013, 2016, 2019, and Online.



## Example

This example below will:

a) Asynchronously collect all events (single and recurring) from the "StaffSchedule" calendar lists. 

b) Return only the events that occur between today and one month from now.

c) Return data from a calendar list that is located at a different subsite than the present one.

    spcalpro.getCalendarEvents({
        listName: 'StaffSchedule',
        getEventsAfterDate: new Date(),
        getEventsBeforeDate: new Date(new Date().getTime() + 2592000000),       // one month from today
        sourceSite: 'https://example.com/Schedules'
    }).ready(function(data, error) {
        if (error) console.error(error);
        console.table(data);
    });
        

Easy enough. Full documentation can be found here: [https://spcalendarpro.sharepointhacks.com](https://spcalendarpro.sharepointhacks.com)
