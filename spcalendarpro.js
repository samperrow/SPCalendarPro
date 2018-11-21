/*
* @name SPCalendarPro
* Version 1.2.1.1
* No dependencies!
* @description An ultra lightweight JavaScript library to easily manage SharePoint calendar events.
* @category Plugins/SPCalendarPro
* @documentation https://spcalendarpro.sharepointhacks.com
* @author Sam Perrow sam.perrow399@gmail.com
*
* Copyright 2018  Sam Perrow  (email : sam.perrow399@gmail.com)
* Licensed under the MIT license:
* http://www.opensource.org/licenses/mit-license.php
*/

(function (global, factory) {
    global.spcalpro = factory();
}(this, function () {

    function getSPEnvInfo(site) {
        var spVersion = _spPageContextInfo.webUIVersion;

        return {
            year: (spVersion === 15) ? '2013' : '2010',
            soapURL: (typeof site === 'string' && site.length > 0)
                ? site + '/_vti_bin/Lists.asmx'
                : (spVersion === 15) ? _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx' : document.location.protocol + '//' + document.location.host + _spPageContextInfo.webServerRelativeUrl + '/_vti_bin/Lists.asmx'
        }
    }

    // checks if supplied datetimes are the same date as ones in calendar list.
    SPCalendarPro.prototype.isSameDate = function () {
        var reqbeginDate = this.userDateTimes.begin.beginDate.toDateString();
        var reqEndDate = this.userDateTimes.end.endDate.toDateString();

        return this.data.filter(function (event) {
            return event.EventDate.toDateString() === reqbeginDate && event.EndDate.toDateString() === reqEndDate;
        });
    }

    // provide begin/end datetimes, and this method will check for events that fall in that range..
    SPCalendarPro.prototype.matchDateTimes = function () {
        var reqBeginDT = this.userDateTimes.begin.beginDateTime;
        var reqEndDT = this.userDateTimes.end.endDateTime;

        return this.data.filter(function (event) {
            return (event.EventDate <= reqBeginDT) && (event.EndDate >= reqEndDT);
        });
    }

    // checks for time conflicts between provided begin/end datetime and events
    SPCalendarPro.prototype.isTimeConflict = function () {
        var reqBeginDT = this.userDateTimes.begin.beginDateTime;
        var reqEndDT = this.userDateTimes.end.endDateTime;

        return this.data.filter(function (event) {
            var arrBeginDT = event.EventDate;
            var arrEndDT = event.EndDate;

            return (
                (reqBeginDT <= arrBeginDT && reqEndDT >= arrEndDT) || (arrBeginDT < reqBeginDT && arrEndDT > reqBeginDT)
                || (arrBeginDT < reqEndDT && arrEndDT > reqEndDT) || (reqBeginDT < arrBeginDT && reqEndDT > arrEndDT));
        });
    }

    // couldn't do without a where clause now could we?
    SPCalendarPro.prototype.where = function (str) {
        var fieldName = str.split(' ')[0];
        var operation = str.split(' ')[1];
        var value = str.split(operation + ' ')[1];

        var operators = {
            '=': function (a, b) { return a == b },
            '>': function (a, b) { return a > new Number(b) },
            '<': function (a, b) { return a < new Number(b) },
            '>=': function (a, b) { return a >= new Number(b) },
            '<=': function (a, b) { return a <= new Number(b) },
            '!=': function (a, b) { return a != b }
        }

        return this.data.filter(function (event) {
            return operators[operation](event[fieldName], value);
        });
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

    function StringToXML(oString) {
        return (window.ActiveXObject)
            ? new ActiveXObject("Microsoft.XMLDOM").loadXML(oString)
            : new DOMParser().parseFromString(oString, 'application/xml');
    }


    // get single, recurring, or all calendar events
    var getListData = function (spCalProObj, userObj, listType, listSourceSite) {
        var doAsync = (typeof userObj.async === 'undefined') ? true : userObj.async;

        // set up the CAML query. returns single and recurring events by default, unless otherwise specified.
        var soapHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" + userObj.listName + "</listName>";
        var soapFooter = "</soap:Body></soap:Envelope>";

        var beginRecurringCaml = "<DateRangesOverlap><FieldRef Name='EventDate'/><FieldRef Name='EndDate'/><FieldRef Name='RecurrenceID'/><Value Type='DateTime'><Year/></Value></DateRangesOverlap>";
        var endRecurringCaml = "</Where><OrderBy><FieldRef Name='EventDate'/></OrderBy></Query></query><queryOptions><QueryOptions><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion><ExpandRecurrence>TRUE</ExpandRecurrence><RecurrenceOrderBy>TRUE</RecurrenceOrderBy><ViewAttributes Scope='RecursiveAll'/></QueryOptions></queryOptions>";

        var singleQuery = (listSourceSite.year === '2010')
            ? "<query><Query><Where><Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>0</Value></Eq></Where></Query></query>"
            : "<query><Query><Where><IsNull><FieldRef Name='fRecurrence'/></IsNull></Where></Query></query>";

        var recurringQuery = "<query><Query><Where><And>" + beginRecurringCaml + "<Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>1</Value></Eq></And>" + endRecurringCaml;
        var query = "";
        var fieldNames = (userObj.fields) ? getFieldNames() : '';

        if (listType === 'calendar') {
            if (userObj.type === 'single') {
                query = singleQuery;
            } else if (userObj.type === 'recurring') {
                query = recurringQuery;
            } else {
                query = "<query><Query><Where>" + beginRecurringCaml + endRecurringCaml;
            }

            query += "<viewFields><ViewFields>" + fieldNames + "</ViewFields></viewFields></GetListItems>";

        } else if (listType === 'list') {

            if (userObj.whereCAMLQuery) {
                query = userObj.whereCAMLQuery;
            }
            query += "<viewFields><ViewFields>" + fieldNames + "</ViewFields></viewFields></GetListItems>";
        }

        function getFieldNames() {
            var viewFields = '<FieldRef Name="fRecurrence"/>';
            for (var i = 0; i < userObj.fields.length; i++) {
                viewFields += "<FieldRef Name='" + userObj.fields[i] + "'/>";
            }
            return viewFields;
        }

        postAjax(soapHeader + query + soapFooter);

        // make ajax request. fires synchronously by default. No j-word needed!
        function postAjax(soapStr) {
            var xhr = new XMLHttpRequest();
            xhr.open('POST', listSourceSite.soapURL, doAsync);
            xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
            xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
            xhr.send(soapStr);

            if (doAsync === true) {
                xhr.onload = function () { 
                    return determineXhrStatus(xhr);
                }
            } else {
                return determineXhrStatus(xhr);
            }
        }

        function determineXhrStatus(xhr) {
            
            if (xhr.readyState == 4 && xhr.status == 200) {
                XhrToObj(xhr);
            } else if (xhr.status == 500) {
                getErrorData(xhr);
            }

            if (userObj.callback) {
                return userObj.callback(spCalProObj.data, spCalProObj);
            } else if (spCalProObj.userCallback) {
                return spCalProObj.userCallback(spCalProObj.data, spCalProObj);
            } else {
                return spCalProObj;
            }

        }

        function XhrToObj(xhr) {
            if (xhr.responseXML) {
                return spCalProObj.data = XmlToJson(xhr.responseXML.querySelectorAll('*'));
            } else if (!xhr.responseXML && xhr.responseText) {                                      // in case the list is an external list. 
                var xml = StringToXML(xhr.responseText.replace(/&#22;|&#0;/g, ''));                 // removes HTML chars that makes the XML parser fail.
                return spCalProObj.data = XmlToJson(xml.querySelectorAll('*'));
            }
        }

        function getErrorData(xhr) {
            spCalProObj.error = {
                errorCode: (/<errorcode xmlns="http:\/\/schemas.microsoft.com\/sharepoint\/soap\/">/.test(xhr.responseText)) ? xhr.responseText.split('<errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorcode>')[0] : '',
                errorString: (/<errorstring xmlns="http:\/\/schemas.microsoft.com\/sharepoint\/soap\/">/.test(xhr.responseText)) ? xhr.responseText.split('<errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorstring>')[0] : '',
                faultString: (/<faultstring>/.test(xhr.responseText)) ? xhr.responseText.split('<faultstring>')[1].split('</faultstring>')[0] : ''
            }
            console.error( spCalProObj.error);
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
                            ? new Date(rowAttrs[attrNum].value.replace('-', '/'))
                            : rowAttrs[attrNum].value;
                    }

                    eventArr.push(row);
                }
            }

            if (userObj.getEventsAfterDate) {
                eventArr = eventArr.filter(function(event) {
                    return event.EventDate >= userObj.getEventsAfterDate;
                });
            }

            if (userObj.getEventsBeforeDate) {
                eventArr = eventArr.filter(function(event) {
                    return event.EventDate <= userObj.getEventsBeforeDate;
                });
            }

            return eventArr;
        }
        return spCalProObj;
    }

    String.prototype.formatInputToHours = function () {
        var amPmTime = this.split(' ');
        var hours = Number(amPmTime[0]);
        return (amPmTime[1] === 'PM' && hours < 12) ? hours += 12 : hours;
    }

    // this will grab date/time input values from a sharepoint form and convert them into proper date objects for later use.
    // by default this grabs the first and second date/time rows from a form.
    var convertFormDateTimes = function (row1, row2) {
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
                time: function () {
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


    // the main object we use.
    function SPCalendarPro(obj, listType) {
        this.listName = obj.listName;
        this.getEventsAfterDate = (obj.getEventsAfterDate) ? obj.getEventsAfterDate : null;
        this.getEventsBeforeDate = (obj.getEventsBeforeDate) ? obj.getEventsBeforeDate : null;
        this.fields = obj.fields ? obj.fields : null;
        this.userDateTimes = (obj.userDateTimes) ? obj.userDateTimes : null;
        this.whereCAMLQuery = (obj.whereCAMLQuery) ? obj.whereCAMLQuery : null;
        var listSourceSite = getSPEnvInfo(obj.sourceSite);

        this.ready = function (execCallback) {
            this.userCallback = execCallback;
        }

        this.callback = function () {
            return (obj.callback) ? obj.callback(this) : null;
        }

        this.data = getListData(this, obj, listType, listSourceSite);
        return this;
    }


    var data = {
        getCalendarEvents: function (obj) {
            return new SPCalendarPro(obj, 'calendar');
        },

        getListItems: function (obj) {
            return new SPCalendarPro(obj, 'list');
        },

        getDateTimesFromForm: function (row1, row2) {
            var time = convertFormDateTimes(row1, row2);
            return formatDateTimesToObj(time.userBeginDT, time.userEndDT);
        },

        provideDateTimes: function (dateTime1, dateTime2) {
            return formatDateTimesToObj(dateTime1, dateTime2);
        },

        disableDragAndDrop: function () {
            ExecuteOrDelayUntilScriptLoaded(disableDragDrop, 'SP.UI.ApplicationPages.Calendar.js');
            function disableDragDrop() {
                var calendarCreate = SP.UI.ApplicationPages.CalendarContainerFactory.create;
                SP.UI.ApplicationPages.CalendarContainerFactory.create = function (elem, cctx, viewType, date, startupData) {
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