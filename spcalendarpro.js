/*
* @name SPCalendarPro
* Version 2018.02
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

<<<<<<< HEAD
        function getSPEnvInfo(obj) {
            var spVersion = _spPageContextInfo.webUIVersion;

            return {
                year: (spVersion === 15) ? '2013' : '2010',
                soapURL: obj.sourceSite 
                    ? obj.sourceSite + '/_vti_bin/Lists.asmx'
                    : (spVersion === 15) ? _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx' : document.location.protocol + '//' + document.location.host + _spPageContextInfo.webServerRelativeUrl + '/_vti_bin/Lists.asmx',
            }
=======
        function getSPEnvInfo() {
            var spVersion = _spPageContextInfo.webUIVersion;
            var versionObj = {};

            if (spVersion === 15) {
                versionObj.year = '2013';
                versionObj.soapURL = _spPageContextInfo.webAbsoluteUrl + '/_vti_bin/Lists.asmx';
                // versionObj.soapURL = "http://sharepoint2013" + '/_vti_bin/Lists.asmx';

            } else {
                versionObj.year = '2010';
                versionObj.soapURL = document.location.protocol + '//' + document.location.host + _spPageContextInfo.webServerRelativeUrl + '/_vti_bin/Lists.asmx';
            }
            return versionObj;
>>>>>>> update
        }

        // checks if supplied datetimes are the same date as ones in calendar list.
        SPCalendarPro.prototype.isSameDate = function() {
            var reqbeginDate = this.userDateTimes.beginDate;
            var reqEndDate = this.userDateTimes.endDate;

            this.listData = this.listData.filter(function(event) {
                return event.EventDate.toDateString() === reqbeginDate && event.EndDate.toDateString() === reqEndDate;
            });

            return this;
        }

        // provide begin/end datetimes, and this method will check for events that fall in that range..
        SPCalendarPro.prototype.matchDateTimes = function() {
            var reqBeginDT = this.userDateTimes.beginDateTime;
            var reqEndDT = this.userDateTimes.endDateTime;

            this.listData = this.listData.filter(function(event) {
                return (event.EventDate <= reqBeginDT) && (event.EndDate >= reqEndDT);
            });

            return this;
        }

        // checks for time conflicts between provided begin/end datetime and events
        SPCalendarPro.prototype.isTimeConflict = function() {
            var reqBeginDT = this.userDateTimes.beginDateTime;
            var reqEndDT = this.userDateTimes.endDateTime;

<<<<<<< HEAD
            if (this.listData) {
                this.listData = this.listData.filter(function(event) {
                    var arrBeginDT = event.EventDate;
                    var arrEndDT = event.EndDate;
=======
            this.events = this.events.filter(function(event) {
                var arrStartDT = event.EventDate;
                var arrEndDT = event.EndDate;
>>>>>>> update
        
                    return (
                        (reqBeginDT <= arrBeginDT && reqEndDT >= arrEndDT) || (arrBeginDT < reqBeginDT && arrEndDT > reqBeginDT)
                        || (arrBeginDT < reqEndDT && arrEndDT > reqEndDT) || (reqBeginDT < arrBeginDT && reqEndDT > arrEndDT) );
                });
            } else {
                console.log( 'no data');
            }


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

            this.listData = this.listData.filter(function(event) {
                return operators[operation](event[fieldName], value);
            });

            return this;
        }

        // this will convert date/time info from a sharepoint form into proper date objects to be used later.
        SPCalendarPro.prototype.getDateTimesFromForm = function(row1, row2) {
            var formattedDT = convertFormDateTimes(row1, row2);
            this.userDateTimes = formatDateTimesToObj(formattedDT.userBeginDT, formattedDT.userEndDT);
            return this;
        }

<<<<<<< HEAD
        // user directly supplies begin and end datetimes
        SPCalendarPro.prototype.getEventsAfterToday = function() {
            this.listData = this.listData.filter(function(event) {
=======
        // // user directly supplies begin and end datetimes
        SPCalendarPro.prototype.getEventsAfterToday = function() {
            this.events = this.events.filter(function(event) {
>>>>>>> update
                var today = new Date();
                return (event.EventDate >= today || event.EndDate >= today);
            });
            return this;
        }

        // // user directly supplies begin and end datetimes
        SPCalendarPro.prototype.provideDateTimes = function(datetime1, datetime2) {
            this.userDateTimes = formatDateTimesToObj(datetime1, datetime2);
            return this;
        }


        // to be used internally, only for formatted the provided datetimes into other formats.
<<<<<<< HEAD
        function formatDateTimesToObj(beginDT, endDT) {
            return {
                beginDateTime: beginDT,
                beginDate: new Date(beginDT.toDateString()),
                beginTime: beginDT.toTimeString(),
=======
        function formatDateTimesToObj(startDT, endDT) {
            return {
                beginDateTime: startDT,
                beginDate: startDT.toDateString(),
                beginTime: startDT.toTimeString(),
>>>>>>> update
                endDateTime: endDT,
                endDate: new Date(endDT.toDateString()),              
                endTime: endDT.toTimeString()
            };
        }

        
        // get single, recurring, or all calendar events
<<<<<<< HEAD
        var getListData = function(spcalproObj, userObj, listType) {
            var doAsync = (typeof userObj.async !== 'undefined') ? userObj.async : true;

            // set up the CAML query. returns single and recurring events by default, unless otherwise specified.
            var soapHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" + userObj.listName + "</listName>";
            var soapFooter = "</soap:Body></soap:Envelope>";

            var beginRecurringCaml = "<DateRangesOverlap><FieldRef Name='EventDate'/><FieldRef Name='EndDate'/><FieldRef Name='RecurrenceID'/><Value Type='DateTime'><Year/></Value></DateRangesOverlap>";
            var endRecurringCaml = "</Where><OrderBy><FieldRef Name='EventDate'/></OrderBy></Query></query><queryOptions><QueryOptions><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion><ExpandRecurrence>TRUE</ExpandRecurrence><RecurrenceOrderBy>TRUE</RecurrenceOrderBy><ViewAttributes Scope='RecursiveAll'/></QueryOptions></queryOptions></GetListItems>";
            var singleQuery = "<query><Query><Where><Eq><FieldRef Name='fRecurrence'/><Value Type='Number'>0</Value></Eq></Where></Query></query></GetListItems>";
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
            } else if (listType === 'list') {
                query = "<viewFields><ViewFields>" + fieldNames + "</ViewFields></viewFields></GetListItems>";
            }
            
            function getFieldNames() {
                var viewFields = '';
                for (var i = 0; i < userObj.fields.length; i++) {
                    viewFields += "<FieldRef Name='" + userObj.fields[i] + "'/>";
                }
                return viewFields;
            }

            postAjax(soapHeader + query + soapFooter);

            // make ajax request. fires synchronously by default. No j-word needed!
            function postAjax(soapStr) {
                var xhr = new XMLHttpRequest();
                xhr.open('POST', spcalproObj.SPInfo.soapURL, doAsync);
                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
                xhr.send(soapStr);
                return (doAsync === true) ? xhr.onload = function() { return getEvents(xhr) } : getEvents(xhr);
            }

            function getEvents(xhr) {
                if (xhr.readyState == 4 && xhr.status == 200) {
                    userObj.listData = XmlToJson( xhr.responseXML.querySelectorAll('*') );
                    if (userObj.callback) return userObj.callback(userObj);
                } else if (xhr.status == 500) {
                    return userObj.error = {
                        errorCode: xhr.responseText.split('<errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorcode>')[0],
                        errorString: xhr.responseText.split('<errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">')[1].split('</errorstring>')[0],
                        faultString: xhr.responseText.split('<faultstring>')[1].split('</faultstring>')[0],
                    }
                }
=======
        var getCalendarEvents = function(obj, async, type) {

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
>>>>>>> update
            }

            var soapStr = soapHeader + query + soapFooter;
            postAjax(soapStr);

            // make ajax request. fires synchronously by default. No j-word needed!
            function postAjax(soapStr) {
                var url = obj.soapUrl;
                var xhr = new XMLHttpRequest();

                xhr.open('POST', url, async);
                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.setRequestHeader('Content-Type', 'text/xml;charset="utf-8"');
                xhr.send(soapStr);

                function getEvents() {
                    if (xhr.readyState == 4 && xhr.status == 200) {
                        obj.events = XmlToJson( xhr.responseXML.querySelectorAll('*') );
                        return (obj.callback) ? obj.callback(obj) : obj;
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
<<<<<<< HEAD
                            var thisAttrName = rowAttrs[attrNum].name.split("ows_")[1];
                                                    
                            row[thisAttrName] = (thisAttrName === 'EventDate' || thisAttrName === 'EndDate')
                                ? new Date(rowAttrs[attrNum].value.replace('-', '/'))
=======
                            var thisAttrName = rowAttrs[attrNum].name;
                            var thisObjectName = thisAttrName.split("ows_")[1];
                            
                            row[thisObjectName] = (thisObjectName === 'EventDate' || thisObjectName === 'EndDate')
                                ? new Date(rowAttrs[attrNum].value.replace('-', '/') )
>>>>>>> update
                                : rowAttrs[attrNum].value;
                        }

<<<<<<< HEAD
                        if (listType === 'calendar' && userObj.getEventsAfterDate) {
                            if (row.EventDate >= userObj.getEventsAfterDate) eventArr.push(row); 
                        } else {
                            eventArr.push(row);
                        }

                    }
                }
                return eventArr;
            }
            return userObj.listData;
=======
            return obj.events;
>>>>>>> update
        }

        String.prototype.formatInputToHours = function() {
            var amPmTime = this.split(' ');
            var hours = Number( amPmTime[0] );
            return (amPmTime[1] === 'PM' && hours < 12) ? hours += 12 : hours;
        }

        // this will grab date/time input values from a sharepoint form and convert them into proper date objects for later use.
        // by default this grabs the first and second date/time rows from a form.
<<<<<<< HEAD
        var iframeContent = document.getElementById('formIframe').contentDocument;

=======
>>>>>>> update
        var convertFormDateTimes = function(row1, row2) {
            row1 = (!row1) ? 0 : row1;
            row2 = (!row2) ? 1 : row2;

            function findDateTimes(row) {
<<<<<<< HEAD
                var dtParentElem = iframeContent.querySelectorAll('input[id$="DateTimeField_DateTimeFieldDate"]')[row].parentNode.parentNode;
=======
                var dtParentElem = document.querySelectorAll('input[id$="DateTimeField_DateTimeFieldDate"]')[row].parentNode.parentNode;
>>>>>>> update
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

<<<<<<< HEAD
            var beginDateTimes = findDateTimes(row1);
            var endDateTimes = findDateTimes(row2);

            return {
                userBeginDT: new Date( beginDateTimes.date + ' ' + beginDateTimes.time() ),
=======
            var startDateTimes = findDateTimes(row1);
            var endDateTimes = findDateTimes(row2);

            return {
                userBeginDT: new Date( startDateTimes.date + ' ' + startDateTimes.time() ),
>>>>>>> update
                userEndDT: new Date( endDateTimes.date + ' ' + endDateTimes.time() )
            }
        }


        // the main object we use.
<<<<<<< HEAD
        function SPCalendarPro(obj, listType) {
            this.listName = obj.listName;
            this.userDateTimes = {};
            this.SPInfo = getSPEnvInfo(obj);
            this.getEventsAfterDate = (obj.getEventsAfterDate) ? obj.getEventsAfterDate : false;
            this.fields = obj.fields ? obj.fields : null;
            
            if (obj.callback) {
                this.callback = function() {
                    return obj.callback(this);
                }
            }
                 
            this.listData = getListData(this, obj, listType);
            return this;
        }


        var data = {
            getCalendarEvents: function(obj) {
                return new SPCalendarPro(obj, 'calendar');
            },
            
            getListItems: function(obj) {
                return new SPCalendarPro(obj, 'list');
=======
        function SPCalendarPro(obj) {
            this.listName = obj.listName;
            this.userDateTimes = {};
            this.soapUrl = (obj.sourceSite) ? obj.sourceSite + '/_vti_bin/Lists.asmx' : getSPEnvInfo().soapURL;

            this.callback = function() {
                return (obj.callback) ? obj.callback(this) : null;
            }

            this.events = getCalendarEvents(this, obj.async, obj.type);
            return this;
        }

        var data = {
            getEvents: function(obj) {
                return new SPCalendarPro(obj);
>>>>>>> update
            },

            getDateTimesFromForm: function(row1, row2) {
                var time = convertFormDateTimes(row1, row2);
                return formatDateTimesToObj( time.userBeginDT, time.userEndDT );
            },
<<<<<<< HEAD

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
            },
=======
>>>>>>> update
            
        }

    return data;

}));