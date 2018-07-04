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
        SPCalendarPro.prototype.getDateTimesFromForm = function(row1 = 0, row2 = 1) {
            var formattedDT = getCalendarEvents(row1, row2);
            return formatDateTimesToObj(this, formattedDT.userBeginDT, formattedDT.userEndDT);
        }

        // user directly supplies begin and end datetimes
        SPCalendarPro.prototype.provideDateTimes = function(datetime1, datetime2) {
            return formatDateTimesToObj(this, datetime1, datetime2);
        }

        // to be used internally, only for formatted the provided datetimes into other formats.
        function formatDateTimesToObj(thisObj, startDT, endDT) {
            thisObj.userDateTimes = {
                startDateTime: startDT,
                startDate: startDT.toDateString(),
                startTime: startDT.toTimeString(),
                endDateTime: endDT,
                endDate: endDT.toDateString(),              
                endTime: endDT.toTimeString()
            };
            return thisObj;
        }

        
        
        // get single, recurring, or all calendar events
        var getCalendarEvents = function(listName, async = true, type, callback) {
            var events = [];

            // set up the CAML query. returns single and recurring events by default, unless otherwise specified.
            var createQuery = function() {
                var startRecurringCaml = "<DateRangesOverlap><FieldRef Name='EventDate'/><FieldRef Name='EndDate'/><FieldRef Name='RecurrenceID'/><Value Type='DateTime'><Year/></Value></DateRangesOverlap>";
                var endRecurringCaml = "</Where><OrderBy><FieldRef Name='EventDate'/></OrderBy></Query></query><queryOptions><QueryOptions><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion><ExpandRecurrence>TRUE</ExpandRecurrence><RecurrenceOrderBy>TRUE</RecurrenceOrderBy><ViewAttributes Scope='RecursiveAll'/></QueryOptions></queryOptions></GetListItems>";
                var query = "";

                //_spPageContextInfo.webUIVersion

                if (type === 'single') {
                    query = "<query><Query><Where><Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>0</Value></Eq></Where></Query></query></GetListItems>";
                } else if (type === 'recurring') {
                    query = "<query><Query><Where><And>" + startRecurringCaml + "<Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>1</Value></Eq></And>" + endRecurringCaml;
                } else {
                    query = "<query><Query><Where>" + startRecurringCaml + endRecurringCaml;
                } 

                return createSoapStr(query);
            }();

            // create the SOAP string we are going to use.
            function createSoapStr(query) {
                var ajaxURL = _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx';
                var soapHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>";
                var list = "<listName>" + listName + "</listName>";
                var soapFooter = "</soap:Body></soap:Envelope>";
                var soapStr = soapHeader + list + query + soapFooter;
                return postAjax(ajaxURL, soapStr, callback);
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

            // make ajax request. fires synchronously by default. No j-word needed!
            function postAjax(url, soapStr) {
                var xhr = new XMLHttpRequest();
                xhr.open('POST', url, async);
                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
                xhr.send(soapStr);
        
                return (async === true) ? xhr.onload = function() { getEvents() } : getEvents();
        
                function getEvents() {
                    if (xhr.readyState == 4 && xhr.status == 200) {
                        return events = XmlToJson( xhr.responseXML.querySelectorAll('*') );
                    }
                }
        
            }  
            
            return events;
        }


        // this will grab date/time input values from a sharepoint form and convert them into proper date objects for later use.
        // by default this grabs the first and second date/time rows from a form.
        var getUserDateTimesFromForm = function(row1 = 0, row2 = 1) {
            var userBeginDT, userEndDT;

            var startDTParentElem = document.querySelectorAll('input[id$="DateTimeField_DateTimeFieldDate"]')[row1].parentNode.parentNode;
            var endDTParentElem = document.querySelectorAll('input[id$="DateTimeField_DateTimeFieldDate"]')[row2].parentNode.parentNode;

            var beginDateElem = startDTParentElem.getElementsByTagName('td')[0].getElementsByTagName('input')[0];
            var beginHours = startDTParentElem.getElementsByClassName('ms-dttimeinput')[0].getElementsByTagName('select')[0];
            var startMin = startDTParentElem.getElementsByClassName('ms-dttimeinput')[0].getElementsByTagName('select')[1];

            var endDateElem = endDTParentElem.getElementsByTagName('td')[0].getElementsByTagName('input')[0];
            var endHours = endDTParentElem.getElementsByClassName('ms-dttimeinput')[0].getElementsByTagName('select')[0];
            var endMin = endDTParentElem.getElementsByClassName('ms-dttimeinput')[0].getElementsByTagName('select')[1];

            String.prototype.formatInputToHours = function() {
                var amPmTime = this.split(' ');
                var hours = Number( amPmTime[0] );
                return (amPmTime[1] === 'PM' && hours < 12) ? hours += 12 : hours;
            }

            var formatDateTimes = function() {
                var beginTime = beginHours.value.formatInputToHours() + ':' + startMin.value;
                var endTime = endHours.value.formatInputToHours() + ':' + endMin.value;
                userBeginDT = new Date( beginDateElem.value + ' ' + beginTime );
                userEndDT = new Date( endDateElem.value + ' ' + endTime );
            }();

            return {
                userBeginDT: userBeginDT,
                userEndDT: userEndDT
            }
        }


        // the main object we use.
        function SPCalendarPro(listName, async, type, callback) {
            this.listName = listName;
            this.userDateTimes = {};

            this.callback = function() {
                console.log( 'hey');
                return callback;
            }

            this.events = getCalendarEvents(this.listName, async, type, callback);


            return this;
        }

        var publicData = {
            getEvents: function(listName, async, type, callback) {
                return new SPCalendarPro(listName, async, type, callback);
            },
        }

    return publicData;

}));