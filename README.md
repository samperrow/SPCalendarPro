# SPCalendarPro

### An ultra lightweight, dependency-free JavaScript library to easily manage SharePoint calendar events.

The painful process of obtaining recurring events, matching user provided datetimes with events, and determining time conflicts is tortorous and requires multiple dependencies and lots of code. SP Calendar Pro simplifies this process enormously.

## Syntax
    spcalpro.getCalendarEvents({
        listName: "StaffSchedule"
    }).ready(function(data, obj) {
        if (obj.error) console.error( obj.error );
        console.table( data );
    });

## Features
1) Easily collects recurring calendar events.
2) Provides a simple way to collect requested datetime values the user.
3) Provides several methods to easily facilitate various datetime comparisons. Match specified datetimes: matchDateTimes(), determine time conflicts: isTimeConflict(), same dates: isSameDate(), etc).
4) A basic 'where' property to filter down the returned data in the CAML query. 
5) Returns regular list items along with calendar events.
6) Option to disable calendar drag and drop: spcalpro.disableDragAndDrop()
7) Requires zero dependencies! Everything is pure vanilla JS.
8) Compatible for all SharePoint versions- 2010, 2013, 2016, 2019, and SP Online.
9) Lightweight! 6 KB minified, 15 KB unminified, with comments.

## Example
The code below will:
1) Asynchronously collect all events (single and recurring) from the `StaffSchedule` calendar list. 
2) Return only the events have a `Title` of `Homer Simpson`.
3) Compare the returned calendar events to see if any pose a time conflict with the datetimes provided in the `.isTimeConflict()` parameters.

## 
    spcalpro.getCalendarEvents({
        listName: "StaffSchedule",
        async: true,
        where: "Title = Homer Simpson"
    }).ready(function(data, obj) {
        if (obj.error) console.error(obj.error);
        var homerJSimpson = obj.isTimeConflict("2019-03-01 00:00:00", "2019-04-01 00:00:00");
        console.table( homerJSimpson );
    });

Full documentation can be found here: [https://sharepointhacks.com/sp-calendar-pro](https://sharepointhacks.com/sp-calendar-pro)
