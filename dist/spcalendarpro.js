/*
 * @name SPCalendarPro
 * Version 1.2.5
 * No dependencies!
 * @description An ultra lightweight JavaScript library to easily manage SharePoint calendar events.
 * @documentation https://sharepointhacks.com/sp-calendar-pro
 * @author Sam Perrow sam.perrow399@gmail.com
 *
 * Copyright 2018  Sam Perrow  (email : sam.perrow399@gmail.com)
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/mit-license.php
 */

(function(global, factory) {
    global.spcalpro = factory();
}(this, function() {

    "use strict";
    var userEnvData;

    function getUserEnvInfo(site) {
        var spVersion = _spPageContextInfo.webUIVersion;

        return {
            year: (spVersion === 15) ? '2013' : '2010',
            soapURL: (typeof site === 'string' && site.length > 0) ?
                site + '/_vti_bin/Lists.asmx' :
                (spVersion === 15) ? _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx' : document.location.protocol + '//' + document.location.host + _spPageContextInfo.webServerRelativeUrl + '/_vti_bin/Lists.asmx'
        }
    }

    // checks if supplied datetimes are the same date as ones in calendar list.
    SPCalendarPro.prototype.isSameDate = function() {
        var reqbeginDate = this.userDateTimes.begin.beginDate;
        var reqEndDate = this.userDateTimes.end.endDate;

        this.data = this.data.filter(function(event) {
            return event.EventDate.toDateString() === reqbeginDate && event.EndDate.toDateString() === reqEndDate;
        });

        return this;
    }

    // provide begin/end datetimes, and this method will check for events that fall in that range..
    SPCalendarPro.prototype.matchDateTimes = function() {
        var reqBeginDT = this.userDateTimes.begin.beginDateTime;
        var reqEndDT = this.userDateTimes.end.endDateTime;

        this.data = this.data.filter(function(event) {
            return (event.EventDate <= reqBeginDT) && (event.EndDate >= reqEndDT);
        });

        return this;
    }

    // checks for time conflicts between provided begin/end datetime and events
    SPCalendarPro.prototype.isTimeConflict = function() {
        var reqBeginDT = this.userDateTimes.begin.beginDateTime;
        var reqEndDT = this.userDateTimes.end.endDateTime;

        this.data = this.data.filter(function(event) {
            var arrBeginDT = event.EventDate;
            var arrEndDT = event.EndDate;

            return (
                (reqBeginDT <= arrBeginDT && reqEndDT >= arrEndDT) || (arrBeginDT < reqBeginDT && arrEndDT > reqBeginDT) ||
                (arrBeginDT < reqEndDT && arrEndDT > reqEndDT) || (reqBeginDT < arrBeginDT && reqEndDT > arrEndDT));
        });

        return this;
    }

    // couldn't do without a where clause now could we?
    SPCalendarPro.prototype.where = function(str) {
        var fieldName = str.split(' ')[0];
        var operation = str.split(' ')[1];
        var value = str.split(operation + ' ')[1];

        var operators = {
            '=': function(a, b) { return a == b },
            '>': function(a, b) { return a > new Number(b) },
            '<': function(a, b) { return a < new Number(b) },
            '>=': function(a, b) { return a >= new Number(b) },
            '<=': function(a, b) { return a <= new Number(b) },
            '!=': function(a, b) { return a != b }
        }

        this.data = this.data.filter(function(event) {
            return operators[operation](event[fieldName], value);
        });

        return this;
    }

    // to be used internally, only for formatted the provided datetimes into other formats.
    function formatDateTimesToObj(beginDT, endDT) {
        return {
            begin: {
                beginDateTime: beginDT,
                beginDate: new Date(beginDT.toDateString()),
                beginTime: beginDT.toTimeString()
            },

            end: {
                endDateTime: endDT,
                endDate: new Date(endDT.toDateString()),
                endTime: endDT.toTimeString()
            }
        };
    }

    // Converts large string from external list to valid XML
    function StringToXML(oString) {
        return (window.ActiveXObject) ?
            new ActiveXObject("Microsoft.XMLDOM").loadXML(oString) :
            new DOMParser().parseFromString(oString, 'application/xml');
    }


    // Query the calendar or list and return the items
    var getListData = function(spCalProObj, userObj, listType, userEnvData) {
        var doAsync = (typeof userObj.async === 'undefined') ? true : userObj.async;



        // Create the CAML query. returns single and recurring events by default, unless otherwise specified.
        function createCAMLQuery() {
            var soapHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" + userObj.listName + "</listName>";
            var soapFooter = "</soap:Body></soap:Envelope>";
            var beginRecurringCaml = "<DateRangesOverlap><FieldRef Name='EventDate'/><FieldRef Name='EndDate'/><FieldRef Name='RecurrenceID'/><Value Type='DateTime'><Year/></Value></DateRangesOverlap>";
            var endRecurringCaml = "</Where><OrderBy><FieldRef Name='EventDate'/></OrderBy></Query></query><queryOptions><QueryOptions><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion><ExpandRecurrence>TRUE</ExpandRecurrence><RecurrenceOrderBy>TRUE</RecurrenceOrderBy><ViewAttributes Scope='RecursiveAll'/></QueryOptions></queryOptions>";
            var recurringQuery = "<query><Query><Where><And>" + beginRecurringCaml + "<Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>1</Value></Eq></And>" + endRecurringCaml;
            var query = "";
            var fieldNames = (userObj.fields) ? getFieldNames() : '';

            if (userObj.CamlQuery) {
                query = userObj.CamlQuery;
            } else if (listType === 'calendar') {
                if (userObj.type === "single") query = "<query><Query><Where><Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>0</Value></Eq></Where></Query></query>";
                else if (userObj.type === 'recurring') query = recurringQuery;
                else query = "<query><Query><Where>" + beginRecurringCaml + endRecurringCaml;
            } else if (listType === 'list' && fieldNames === "") {
                query = "<viewFields><ViewFields></ViewFields></viewFields>";
            }

            query += fieldNames + "</GetListItems>";
            postAjax(soapHeader + query + soapFooter);
        }
        createCAMLQuery();


        function getFieldNames() {
            var viewFields = (listType === 'calendar') ? '<FieldRef Name="fRecurrence"/>' : '';

            for (var i = 0; i < userObj.fields.length; i++) {
                if (typeof userObj.fields[i] === "string") {
                    viewFields += "<FieldRef Name='" + userObj.fields[i] + "'/>";
                }
            }

            return (viewFields.length > 0) ? "<viewFields><ViewFields>" + viewFields + "</ViewFields></viewFields>" : '';
        }


        // make ajax request. fires synchronously by default.
        function postAjax(soapStr) {
            var xhr = new XMLHttpRequest();
            xhr.open('POST', userEnvData.soapURL, doAsync);
            xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
            xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
            xhr.send(soapStr);

            if (doAsync === true) {
                xhr[(xhr.onload) ? "onload" : "onreadystatechange"] = function() {
                    return determineXhrStatus(xhr);
                }
            } else {
                determineXhrStatus(xhr);
            }
        }

        function determineXhrStatus(xhr) {
            if (xhr.readyState == 4 && xhr.status == 200) {
                XhrToObj(xhr);
                complete();
            } else if (xhr.status == 500) {
                getErrorData(xhr);
                complete();
            }
        }

        function XhrToObj(xhr) {
            if (xhr.responseXML) {
                return spCalProObj.data = XmlToJson(xhr.responseXML.getElementsByTagName('*'));
            } else if (!xhr.responseXML && xhr.responseText) { // in case the list is an external list. 
                var xml = StringToXML(xhr.responseText.replace(/&#22;|&#0;/g, '')); // removes HTML chars that makes the XML parsing fail.
                return spCalProObj.data = XmlToJson(xml.getElementsByTagName('*'));
            }
        }

        function getErrorData(xhr) {
            spCalProObj.error = {
                errorCode: (/<errorcode xmlns="http:\/\/schemas.microsoft.com\/sharepoint\/soap\/">/.test(xhr.responseText)) ? xhr.responseText.split('<errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorcode>')[0] : '',
                errorString: (/<errorstring xmlns="http:\/\/schemas.microsoft.com\/sharepoint\/soap\/">/.test(xhr.responseText)) ? xhr.responseText.split('<errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorstring>')[0] : '',
                faultString: (/<faultstring>/.test(xhr.responseText)) ? xhr.responseText.split('<faultstring>')[1].split('</faultstring>')[0] : ''
            }
            console.error(spCalProObj.error);
            return spCalProObj.error;
        }

        // accepts XML, returns an array of objects, each of which are calendar events.
        function XmlToJson(xml) {
            var eventArr = [];

            for (var i = 0; i < xml.length; i++) {
                var row = {};
                var rowAttrs = xml[i].attributes;

                if (xml[i].nodeName === 'z:row') {

                    for (var attrNum = 0; attrNum < rowAttrs.length; attrNum++) {
                        var thisAttrName = rowAttrs[attrNum].name.split("ows_")[1];

                        row[thisAttrName] = (thisAttrName === 'EventDate' || thisAttrName === 'EndDate') 
                            ? new Date(rowAttrs[attrNum].value.replace(/-/g, '/')) 
                            : rowAttrs[attrNum].value;
                    }

                    eventArr.push(row);
                }
            }

            if (spCalProObj.getEventsAfterDate) {
                eventArr = eventArr.filter(function(event) {
                    return (event.EventDate) ? event.EventDate >= spCalProObj.getEventsAfterDate : event;
                });
            }

            if (spCalProObj.getEventsBeforeDate) {
                eventArr = eventArr.filter(function(event) {
                    return (event.EventDate) ? event.EventDate <= spCalProObj.getEventsBeforeDate : event;
                });
            }

            return eventArr;
        }

        function complete() {
            if (userObj.callback) {
                return userObj.callback(spCalProObj.data, spCalProObj);
            } else if (spCalProObj.userCallback) {
                return spCalProObj.userCallback(spCalProObj.data, spCalProObj);
            } else {
                return spCalProObj;
            }
        }

        return spCalProObj.data;
    }

    String.prototype.formatInputToHours = function() {
        var amPmTime = this.split(' ');
        var hours = Number(amPmTime[0]);
        return (amPmTime[1] === 'PM' && hours < 12) ? hours += 12 : hours;
    }

    // this will grab date/time input values from a sharepoint form and convert them into proper date objects for later use.
    // by default this grabs the first and second date/time rows from a form.
    var convertFormDateTimes = function(row1, row2) {
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
                }
            }
        }

        var beginDateTimes = findDateTimes(row1);
        var endDateTimes = findDateTimes(row2);

        return {
            userBeginDT: new Date(beginDateTimes.date + ' ' + beginDateTimes.time()),
            userEndDT: new Date(endDateTimes.date + ' ' + endDateTimes.time())
        }
    }

    // taken from https://stackoverflow.com/questions/7153470/why-wont-filter-work-in-internet-explorer-8
    function createArrayFilter() {
        return Array.prototype.filter = function(fn) {

            if (this === void 0 || this === null)
                throw new TypeError();

            var t = Object(this);
            var len = t.length >>> 0;
            if (typeof fn !== "function")
                throw new TypeError();

            var res = [];
            var thisp = arguments[1];
            for (var i = 0; i < len; i++) {
                if (i in t) {
                    var val = t[i]; // in case fn mutates this
                    if (fn.call(thisp, val, i, t)) {
                        res.push(val);
                    }
                }
            }

            return res;
        }
    }


    // Turn the user provided value into a date object if needed
    function checkDateType(val) {
        return (val) ? (typeof val.getMonth === "function") ? val : new Date(val) : null;
    }


    // the main object we use.
    function SPCalendarPro(obj, listType) {
        this.listName = (obj.listName) ? obj.listName : null;
        this.getEventsAfterDate = checkDateType(obj.getEventsAfterDate);
        this.getEventsBeforeDate = checkDateType(obj.getEventsBeforeDate);
        this.fields = obj.fields ? obj.fields : null;
        this.userDateTimes = (obj.userDateTimes) ? obj.userDateTimes : null;
        this.CamlQuery = (obj.CamlQuery) ? obj.CamlQuery : null;
        userEnvData = getUserEnvInfo(obj.sourceSite);

        if (!Array.prototype.filter) {
            createArrayFilter();
        }

        this.ready = function(execCallback) {
            this.userCallback = execCallback;
        }

        this.callback = function() {
            return (obj.callback) ? obj.callback(this) : null;
        }

        if (typeof obj.listName === "string") {
            this.data = getListData(this, obj, listType, userEnvData);
        } else {
            console.error('You must specify a list name.');
        }
        return this;
    }


    var data = {
        getCalendarEvents: function(obj) {
            return new SPCalendarPro(obj, 'calendar');
        },

        getListItems: function(obj) {
            return new SPCalendarPro(obj, 'list');
        },

        getDateTimesFromForm: function(row1, row2) {
            var time = convertFormDateTimes(row1, row2);
            return formatDateTimesToObj(time.userBeginDT, time.userEndDT);
        },

        userDates: function(dateTime1, dateTime2) {
            return formatDateTimesToObj(checkDateType(dateTime1), checkDateType(dateTime2));
        },

        disableDragAndDrop: function() {
            ExecuteOrDelayUntilScriptLoaded(disableDragDrop, 'SP.UI.ApplicationPages.Calendar.js');

            function disableDragDrop() {
                var calendarCreate = SP.UI.ApplicationPages.CalendarContainerFactory.create;
                SP.UI.ApplicationPages.CalendarContainerFactory.create = function(elem, cctx, viewType, date, startupData) {
                    if (cctx.dataSources && cctx.dataSources instanceof Array && cctx.dataSources.length > 0) {
                        for (var i = 0; i < cctx.dataSources.length; i++) {
                            cctx.dataSources[i].disableDrag = true;
                        }
                    }
                    calendarCreate(elem, cctx, viewType, date, startupData);
                }
            }
        }

    }

    return data;

}));