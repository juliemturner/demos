"use strict";
var C3C = window.C3C || {};
C3C.currentSite = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;

C3C.init = function () {
    //Global Variables, List name, requests array, define items object
    var REQUESTS_LIST = "IT Requests";
    var requests = [];
    var itemObj = function() {
            return {ID: null, BusinessUnit: null, Category: null, Status: null, DueDate: null, Assigned: null};
        };

    //Internal lookup function to determine if assigned user and business unit combination already have been added to the passed array.
    function getValueAssigned(lookupArray, bu, assigned) {
        var retVal = 0;
        $.each(lookupArray, function () {
            if (this.assigned == assigned && this.businessunit == bu) {
                retVal = this.assignments;
                return false;
            }
        });
        return retVal;
    }

    //Load requests function that make async REST call to get SharePoint data from list.
    var loadRequests = function() {
        $.ajax({
            url: C3C.currentSite + "/_api/web/lists/getbytitle('" + REQUESTS_LIST + "')/items?$top=5000&$select=ID,BusinessUnit,Category,Status,DueDate,AssignedTo/Title&$expand=AssignedTo/Title&$filter=(Status eq 'New') or (Status eq 'Active')",
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            success: loadRequestsSuccess,
            error: error
        });
    };

    //Function that acts on data returned from REST call
    var loadRequestsSuccess = function (itResponse) {
        requests = [];

        var it = itResponse.d.results;
        //Loop through list items, massage data and build new array of data in requests []
        for (var i = 0; i < it.length; i++) {
            var item = itemObj();
            item.ID = it[i].ID;
            item.BusinessUnit = it[i].BusinessUnit;
            item.Category = it[i].Category;
            item.Status = it[i].Status;
            if(it[i].DueDate != undefined)
                item.DueDate = new Date(it[i].DueDate);
            if(it[i].AssignedTo != undefined)
                item.Assigned = it[i].AssignedTo.Title.split(" ")[0];
            requests.push(item);
        }
        
        //The data has been massaged, now format it so that the charting library can understand it.
        chartRequestsByAssignee();
    };

    //Massage data into format charting library can understsand.
    var chartRequestsByAssignee = function () {
        //Local variable to hold an array of [Assigned, BusinessUnit, Assignments]
        var assignedrequests = [];
        $.each(requests, function () {
            if (this.Assigned != null) {
                var i = -1;
                var assigned = this.Assigned;
                var bu = this.BusinessUnit;
                $.each(assignedrequests, function (index) {
                    if (this.assigned == assigned && this.businessunit == bu) { 
                        i = index;
                        return false;
                    }
                });
                if (i == -1)
                    assignedrequests.push({businessunit: bu, assigned: assigned, assignments: 1 });  
                else
                    assignedrequests[i].assignments++;
            }
        });

        var cat = [];
        var val = [];
        //Create simple string array of unique business units. [businessunit]
        $.each(assignedrequests, function () {
            var bu = this.businessunit;
            if ($.inArray(bu, cat) == -1) {
                cat.push(bu);
            }
        });
        var a = [];
        //Create simple string array of unique Assignees [username]
        $.each(assignedrequests, function () {
            var assigned = this.assigned;
            if ($.inArray(assigned, a) == -1) {
                a.push(assigned);
            }
        });
        //Create values array of assignee objects with their assignments for each business unit [[Assignee, BU1-assignments, BU2-assignments, etc]]
        //eg.. [['Jack',1,4,3,2,5],['Jill',6,3,7,3,5]]
        $.each(a, function () {
            var assigned = this.toString();
            var dataArray = [];
            dataArray.push(assigned);
            $.each(cat, function () {
                var bu = this;
                dataArray.push(getValueAssigned(assignedrequests, bu, assigned));
            });
            val.push(dataArray);
        });
        //Push an 'x' as the first element of the business unit array, this is how c3 determines that these are categories for the area chart
        cat.splice(0,0,'x');
        //Push this 'categories' array onto our 'values' array as the first element.
        //eg.. [['x','BU1','BU2','BU3','BU4','BU5'],['Jack',1,4,3,2,5],['Jill',6,3,7,3,5]]
        val.splice(0,0,cat);

        var Title = "Requests by Assignee";
        //Make call to c3.generate command, passing object with configuration settings and values array.
        var reqByAssignee = c3.generate({
            bindto: '#chart',
            data: {
                x: 'x',
                columns: val,
                type: 'area'
            },
            axis: {
                x: {
                    type: 'category'
                }
            }
        });
    };

    //Error function if REST call throws error
    var error = function (sender, args) {
        alert(args.get_message());
    };

    loadRequests();
};

C3C.init();