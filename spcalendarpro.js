/*
* @name SPCalendarPro
* Version 2018.01
* No dependencies!
* @description An ultra lightweight JavaScript library to easily manage SharePoint calendar events.
* @category Plugins/SPCalendarPro
* @author Sam Perrow sam.perrow399@gmail.com
*
* Copyright 2018  Sam Perrow  (email : sam.perrow399@gmail.com)
* Licensed under the MIT license:
* http://www.opensource.org/licenses/mit-license.php
*/

(function (global, factory) {
    global.spcalpro = factory();
    }(this, function() {

        // checks if supplied datetimes are the same date as ones in calendar list.
        SPCalendarPro.prototype.isSameDate = function() {
            var reqStartDate = this.userDateTimes.startDate;
            var reqEndDate = this.userDateTimes.endDate;

            this.events = this.events.filter(function(event) {
                return new Date(event.EventDate).toDateString() === reqStartDate && new Date(event.EndDate).toDateString() === reqEndDate;
            });

            return this;
        }

        // provide begin/end datetimes, and this method will check for events that fall in that range..
        SPCalendarPro.prototype.matchDateTimes = function() {
            var reqStartDT = this.userDateTimes.startDateTime;
            var reqEndDT = this.userDateTimes.endDateTime;

            this.events = this.events.filter(function(event) {
                return (new Date(event.EventDate) <= reqStartDT) && (new Date(event.EndDate) >= reqEndDT);
            });

            return this;
        }

        // checks for time conflicts between provided begin/end datetime and events
        SPCalendarPro.prototype.isTimeConflict = function() {
            var reqStartDT = this.userDateTimes.startDateTime;
            var reqEndDT = this.userDateTimes.endDateTime;

            this.events = this.events.filter(function(event) {
                var arrStartDT = new Date(event.EventDate);
                var arrEndDT = new Date(event.EndDate);
        
                return (
                    (reqStartDT <= arrStartDT && reqEndDT >= arrEndDT) || (arrStartDT < reqStartDT && arrEndDT > reqStartDT)
                    || (arrStartDT < reqEndDT && arrEndDT > reqEndDT) || (reqStartDT < arrStartDT && reqEndDT > arrEndDT) );
            });

            return this;
        }

        // couldn't do without a where clause now could we?
        SPCalendarPro.prototype.where = function(str) {
            var fieldName = str.split(' ')[0];
            var operation = str.split(' ')[1];
            var value = str.split(' ')[2];

            var operators = {
                '=': function(a, b) { return a == b },
                '>': function(a, b) { return a > new Number(b) },
                '<': function(a, b) { return a < new Number(b) },
                '>=': function(a, b) { return a >= new Number(b) },
                '<=': function(a, b) { return a <= new Number(b) },
                '!=': function(a, b) { return a != b },
            }

            this.events = this.events.filter(function(event) {
                return operators[operation](event[fieldName], value);
            });

            return this;
        }

        // this will convert date/time info from a sharepoint form into proper date objects to be used later.
        SPCalendarPro.prototype.getDateTimesFromForm = function(row1, row2) {
            var formattedDT = getDateTimesFromForm(row1, row2);
            return formatDateTimesToObj(this, formattedDT.userBeginDT, formattedDT.userEndDT);
        }

        // user directly supplies begin and end datetimes
        SPCalendarPro.prototype.provideDateTimes = function(datetime1, datetime2) {
            return formatDateTimesToObj(this, datetime1, datetime2);
        }

        // to be used internally, only for formatted the provided datetimes into other formats.
        function formatDateTimesToObj(thisObj, startDT, endDT) {
            thisObj.userDateTimes = {
                beginDateTime: startDT,
                beginDate: startDT.toDateString(),
                beginTime: startDT.toTimeString(),
                endDateTime: endDT,
                endDate: endDT.toDateString(),              
                endTime: endDT.toTimeString()
            };
            return thisObj;
        }

        
        // get single, recurring, or all calendar events
        var getCalendarEvents = function(obj, async, type, cb) {

            // set up the CAML query. returns single and recurring events by default, unless otherwise specified.
            var soapHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" + obj.listName + "</listName>";
            var soapFooter = "</soap:Body></soap:Envelope>";

            var startRecurringCaml = "<DateRangesOverlap><FieldRef Name='EventDate'/><FieldRef Name='EndDate'/><FieldRef Name='RecurrenceID'/><Value Type='DateTime'><Year/></Value></DateRangesOverlap>";
            var endRecurringCaml = "</Where><OrderBy><FieldRef Name='EventDate'/></OrderBy></Query></query><queryOptions><QueryOptions><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion><ExpandRecurrence>TRUE</ExpandRecurrence><RecurrenceOrderBy>TRUE</RecurrenceOrderBy><ViewAttributes Scope='RecursiveAll'/></QueryOptions></queryOptions></GetListItems>";
            var query = "";

            if (type === 'single') {
                query = "<query><Query><Where><Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>0</Value></Eq></Where></Query></query></GetListItems>";
            } else if (type === 'recurring') {
                query = "<query><Query><Where><And>" + startRecurringCaml + "<Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>1</Value></Eq></And>" + endRecurringCaml;
            } else {
                query = "<query><Query><Where>" + startRecurringCaml + endRecurringCaml;
            }

            var soapStr = soapHeader + query + soapFooter;
            postAjax(soapStr);

            // make ajax request. fires synchronously by default. No j-word needed!
            function postAjax(soapStr) {
                var url = _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx';
                var xhr = new XMLHttpRequest();

                xhr.open('POST', url, async);
                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
                xhr.send(soapStr);

                function getEvents() {
                    if (xhr.readyState == 4 && xhr.status == 200) {
                        obj.events = XmlToJson( xhr.responseXML.querySelectorAll('*') );
                        return obj.callback(obj);
                    }
                }

                return (async === true) ? xhr.onload = function() { getEvents() } : getEvents();
            }    

            // accepts XML, returns an array of objects, each of which are calendar events.
            function XmlToJson(xml) {
                var eventArr = [];

                for (var i = 0; i < xml.length; i++) {
                    var row = {};
                    var rowAttrs = xml[i].attributes;

                    if (xml[i].nodeName === 'z:row') {
                        for (var attrNum = 0; attrNum < rowAttrs.length; attrNum++) {
                            var thisAttrName = rowAttrs[attrNum].name;
                            var thisObjectName = thisAttrName.split("ows_")[1];
                            row[thisObjectName] = rowAttrs[attrNum].value;
                        }
                        eventArr.push(row);
                    }
                }
                return eventArr;
            }

            return obj.events;
        }

        String.prototype.formatInputToHours = function() {
            var amPmTime = this.split(' ');
            var hours = Number( amPmTime[0] );
            return (amPmTime[1] === 'PM' && hours < 12) ? hours += 12 : hours;
        }

        // this will grab date/time input values from a sharepoint form and convert them into proper date objects for later use.
        // by default this grabs the first and second date/time rows from a form.
        var getDateTimesFromForm = function(row1, row2) {
            row1 = (!row1) ? 0 : row1;
            row2 = (!row2) ? 1 : row2;

            function findDateTimes(row) {
                var dtParentElem = document.querySelectorAll('input[id$="DateTimeField_DateTimeFieldDate"]')[row].parentNode.parentNode;
                var timeElem = dtParentElem.getElementsByClassName('ms-dttimeinput')[0];
                
                if (timeElem) {
                    var hours = timeElem.getElementsByTagName('select')[0].value;
                    var min = timeElem.getElementsByTagName('select')[1].value;
                }

                return {
                    date: dtParentElem.getElementsByTagName('td')[0].getElementsByTagName('input')[0].value,
                    time: function() {
                        return (hours && min) ? hours.formatInputToHours() + ':' + min : '';
                    },
                }
            }

            var startDateTimes = findDateTimes(row1);
            var endDateTimes = findDateTimes(row2);

            return {
                userBeginDT: new Date( startDateTimes.date + ' ' + startDateTimes.time() ),
                userEndDT: new Date( endDateTimes.date + ' ' + endDateTimes.time() )
            }
        }


        // the main object we use.
        function SPCalendarPro(listName, async, type, cb) {
            this.listName = listName;
            this.userDateTimes = {};

            if (cb) {
                this.callback = function() {
                    return cb(this);
                }
            }

            this.events = getCalendarEvents(this, async, type, cb);
            return this;
        }

        var data = {
            getEvents: function(listName, async, type, cb) {
                return new SPCalendarPro(listName, async, type, cb);
            },
        }

    return data;

}));
               