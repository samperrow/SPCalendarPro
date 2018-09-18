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

a) Asynchronously collect all events (single and recurring) from the "StaffSchedule", "SurgeonSchedule", and the "Appointments" calendar lists. 

c) Convert user provided datetime information from a SharePoint form into proper date objects for use. The "0,1" parameters specify which datetime field elements on a form are to be converted.

d) Determine if user provided datetime values pose a time conflict with any events from the Appointments list.

e) Determine if there user provided date time values coincide with availabilites in the "StaffSchedule" and the "SurgeonSchedule" calendars.

f) The above four steps are executed in the callback method, which is checkApptAvailability().

    var patientTimes = spcalpro.getDateTimesFromForm(0,1);   // Convert the start and end datetimes on a SharePoint form into valid JavaScript dates.

    var staffSchedule = spcalpro.getCalendarEvents({
        listName: "StaffSchedule",
        userDateTimes: patientTimes,
        callback: checkApptAvailability
    });

    var surgeonSchedule = spcalpro.getCalendarEvents({
        listName: "SurgeonSchedule",
        userDateTimes: patientTimes,
        callback: checkApptAvailability
    });

    var appointments = spcalpro.getCalendarEvents({
        listName: "Appointments",
        type: "single",                                 // Returns only single events.
        userDateTimes: patientTimes,
        callback: checkApptAvailability
    });

    function checkApptAvailability() {

        // Ensure all calendar data has been collected before we filter the events.
        if ( staffSchedule.listData && surgeonSchedule.listData && appointments.listData) {
            staffSchedule.matchDateTimes();
            surgeonSchedule.matchDateTime();
            appointments.isTimeConflict();

            if ( staffSchedule.listData.length > 0 && surgeonSchedule.listData.length > 0 && appointments.listData.length < 1 ) {
                return confirmAppt(true);
            }
        }
    }

    function confirmAppt(status) {
        alert("Your appointment has been confirmed!");
    }
    

Easy enough. Full documentation can be found here: [https://spcalendarpro.sharepointhacks.com](https://spcalendarpro.sharepointhacks.com)
