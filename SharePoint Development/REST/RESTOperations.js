/*
This file provides all operations available in SharePoint REST Api
AUTHOR :- FEBIN JS
*/
var RESTOperations = window.OPMNameSpace || {};

var RESTOperations ={
  this.Methods = {
    /*List Operations*/

    /*List Item Operations*/
    CreateItem = function(restURL, data, success, failure){
      $.ajax({
          url: restURL,
          type: "POST",
          async: false,
          contentType: "application/json;odata=verbose",
          data: JSON.stringify(data),
          headers: {
              "Accept": "application/json;odata=verbose",
              "X-RequestDigest": $("#__REQUESTDIGEST").val()
          },
          success: function (data) {
              success(data);
          },
          error: function (error) {
              failure(error);
          }
      });
    },
    UpdateItem = function(restURL, data, success, failure){
      $.ajax({
        url: restURL,
        type: "POST",
        async: false,
        data: JSON.stringify(data),
        contentType: "application/json;odata=verbose",
        headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "X-HTTP-Method": "MERGE",
            "If-Match": "*"
        },
        success: function (data) {
            success(data);
        },
        error: function (error) {
            failure(error);
        }
      });
    },
    RetriveItem = function(restURL, success, failure){
      $.ajax({
        url: restURL,
        type: "GET",
        async: false,
        headers: {
            "accept": "application/json;odata=verbose",
        },
        success: function (data) {
            success(data);
        },
        error: function (error) {
            failure(error);
        }
      });
    },
    DeleteItem = function(restURL, success, failure){
        $.ajax({
         url: restURL,
         type: "DELETE",
         headers: {
           "Accept": "application/json; odata=verbose",
           "X-RequestDigest": $("#__REQUESTDIGEST").val(),
           "If-Match": "*"
         },
         success: success(data),
         error:failure(error)
      });
    },
    GetItemTypeForListName = function(listName) {
     var listItemType;
     $.ajax({
         url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/ListItemEntityTypeFullName",
         type: "GET",
         async: false,
         contentType: "application/json;odata=verbose",
         headers: {
             "Accept": "application/json;odata=verbose",
             "X-RequestDigest": $("#__REQUESTDIGEST").val()
         },
         success: function (data) {
             listItemType = data.d.ListItemEntityTypeFullName;

         },
         error: function (error) {
             console.log("Error" + JSON.stringify(error));
         }
     });
     return listItemType;
    }
    /*Document Library Operations*/
    /*Field Operations*/
 }
}
