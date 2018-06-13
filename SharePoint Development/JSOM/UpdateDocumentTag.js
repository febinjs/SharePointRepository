/*Below is the format of Array Of Objects needs to be created before using the StartOperations()*/
var documentArray = [
		{ ID: 1, ServiceLine: "AM Doc Mgmt|5517ec78-d74c-4ef5-b3a6-63c74e73345e", Status: "Approved|773fc1d6-657f-4260-8bba-5963c95c0dc6" },
		{ ID: 2, ServiceLine: "AM Other Items|ea4d5ac9-e77c-4015-b50f-9bcc6fc8a2fe", Status: "Pending|1cbaee48-3816-4fef-be5b-978a8950815b" },
		{ ID: 3, ServiceLine: "AM Tracking|9667e278-15be-4f5c-89e8-3293cb70baf4", Status: "Rejected|542acb3a-12df-46da-a238-d3d9eec280c8" },
		{ ID: 4, ServiceLine: "AM Doc Mgmt|5517ec78-d74c-4ef5-b3a6-63c74e73345e", Status: "Approved|773fc1d6-657f-4260-8bba-5963c95c0dc6" },
		{ ID: 5, ServiceLine: "AM Other Items|ea4d5ac9-e77c-4015-b50f-9bcc6fc8a2fe", Status: "Pending|1cbaee48-3816-4fef-be5b-978a8950815b" },
		{ ID: 6, ServiceLine: "AM Tracking|9667e278-15be-4f5c-89e8-3293cb70baf4", Status: "Rejected|542acb3a-12df-46da-a238-d3d9eec280c8" },
		{ ID: 7, ServiceLine: "AM Doc Mgmt|5517ec78-d74c-4ef5-b3a6-63c74e73345e", Status: "Approved|773fc1d6-657f-4260-8bba-5963c95c0dc6" },
		{ ID: 8, ServiceLine: "AM Other Items|ea4d5ac9-e77c-4015-b50f-9bcc6fc8a2fe", Status: "Pending|1cbaee48-3816-4fef-be5b-978a8950815b" },
		{ ID: 9, ServiceLine: "AM Tracking|9667e278-15be-4f5c-89e8-3293cb70baf4", Status: "Rejected|542acb3a-12df-46da-a238-d3d9eec280c8" },
		{ ID: 10, ServiceLine: "AM Doc Mgmt|5517ec78-d74c-4ef5-b3a6-63c74e73345e", Status: "Approved|773fc1d6-657f-4260-8bba-5963c95c0dc6" },
		{ ID: 11, ServiceLine: "AM Other Items|ea4d5ac9-e77c-4015-b50f-9bcc6fc8a2fe", Status: "Pending|1cbaee48-3816-4fef-be5b-978a8950815b" },
		{ ID: 12, ServiceLine: "AM Tracking|9667e278-15be-4f5c-89e8-3293cb70baf4", Status: "Rejected|542acb3a-12df-46da-a238-d3d9eec280c8" },
		{ ID: 13, ServiceLine: "" },
		{ ID: 14, ServiceLine: "AM Doc Mgmt|5517ec78-d74c-4ef5-b3a6-63c74e73345e", Status: "Approved|773fc1d6-657f-4260-8bba-5963c95c0dc6" }
];


var oList = null;
var clientContext = null;
var currentListName = "";
/*Array of MMS Field Internal names which is to be supplied*/
var metadataFieldConfigArray = ["ServiceLine", "Status"];
var finalDocArray = [];
var cntr;

/* Below Function Loads the Taxonomy JS File */
function StartOperations() {
    metadataFieldConfigArray = metadataFieldConfigArray;
    var scriptbase = _spPageContextInfo.webServerRelativeUrl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.Runtime.js", function () {
        $.getScript(scriptbase + "SP.js", function () {
            $.getScript(scriptbase + "SP.Taxonomy.js", function () {
                GetListDetails().then(function () {
                    cntr = 0;
                    finalDocArray = chunkArray(documentArray, 10);
                    UpdateBulkItems();
                },
                function (error) {
                    console.log('Metadata Updation failed: ' + error);
                });
            });
        });
    }
	);
}

/* Below Function Loads the SharePoint List */
function GetListDetails() {
    var deferred = $.Deferred();
    clientContext = SP.ClientContext.get_current();

    /* Get List - Change List Name */
    oList = clientContext.get_web().get_lists().getByTitle(currentListName);

    clientContext.executeQueryAsync(function () {
        deferred.resolve();
    }, function (sender, args) {
        deferred.reject(args.get_message());
    });

    return deferred.promise();
}

/* Below Function updates metadata in the Library */
function BulkUpdateItems(docArray) {
    var deferred = $.Deferred();
    var itemArray = [];
    for (var i = 0; i < docArray.length; i++) {
        var documentID = docArray[i].ID;
        var oListItem = oList.getItemById(documentID);
        var isEmpty = false;

        for (var j = 0; j < metadataFieldConfigArray.length; j++) {

            var mmsValue = docArray[i][metadataFieldConfigArray[j]];

            if (mmsValue != "" && mmsValue != null) {

                var field = oList.get_fields().getByInternalNameOrTitle(metadataFieldConfigArray[j]);
                var taxField = clientContext.castTo(field, SP.Taxonomy.TaxonomyField);

                var mmsValArray = mmsValue.split('|');
                var termValue = new SP.Taxonomy.TaxonomyFieldValue();
                termValue.set_label(mmsValArray[0]);
                termValue.set_termGuid(mmsValArray[1]);
                termValue.set_wssId(-1);
                taxField.setFieldValueByValue(oListItem, termValue);
                clientContext.load(taxField);
            } else {
                isEmpty = true;
            }

        }

        if (!isEmpty) {
            oListItem.update();
            itemArray[i] = oListItem;

            clientContext.load(itemArray[i]);
        }
    }

    clientContext.executeQueryAsync(function (sender, args) {
        deferred.resolve();
    }, function (sender, args) {
        deferred.reject(args.get_message());
    });
    return deferred.promise();
}

/*Loop through the input array of data and performs batching operation*/
function UpdateBulkItems() {
    BulkUpdateItems(finalDocArray[cntr])
	.then(
		function () {
		    cntr++;
		    if (cntr < finalDocArray.length) {
		        UpdateBulkItems();
		    }
		    else {
		        alert('Metadata Updated Successfully');
		    }
		},
		function (error) {
		    console.log('Metadata Updation failed: ' + error);
		}
	);
}

/*Returns an array with arrays of the given size.
 * @param myArray {Array} Array to split
 * @param chunkSize {Integer} Size of every group*/
function chunkArray(myArray, chunkSize) {
    var results = [];

    while (myArray.length) {
        results.push(myArray.splice(0, chunkSize));
    }

    return results;
}

/*Below line Invokes the StartOperations()*/
StartOperations();
