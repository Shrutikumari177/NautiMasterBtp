sap.ui.define(
    [
      "sap/ui/core/mvc/Controller",
      "sap/ui/core/Fragment",
      "sap/ui/core/routing/History",
      "sap/m/MessageToast",
      "sap/ui/export/Spreadsheet",
    ],
    /**
     * @param {typeof sap.ui.core.mvc.Controller} Controller
     */
    function (Controller, Fragment, History, MessageToast, Spreadsheet) {
      "use strict";
   
      return Controller.extend("nauticalfe.controller.MastCountryMasterUpload", {
        onInit: function () {},
       
        onBackPress: function () {
          const oRouter = this.getOwnerComponent().getRouter();
          const oFileUploader = this.getView().byId("fileUploader");
          oFileUploader.clear();
          oRouter.navTo("RouteMasterDashboard");
        },
        onPressHome: function () {
          const oRouter = this.getOwnerComponent().getRouter();
          oRouter.navTo("RouteHome");
        },
       
        onUploadPress: function () {
          var fileUploader = this.getView().byId("fileUploader");
          var file = fileUploader.oFileUpload.files[0];
   
          // Check if a file is selected
          if (!file) {
            // Show an error message to the user using MessageToast
            sap.m.MessageToast.show("Please select a file to upload.");
            return; // Exit the function
          }
   
          // Proceed with file upload logic
          fileUploader.upload();
        },
   
        onUploadComplete: function (oEvent) {
          var fileUploader = this.getView().byId("fileUploader");
          var file = fileUploader.oFileUpload.files[0];
   
          // Check if a file is selected
          if (!file) {
            // Show an error message to the user using MessageToast
            sap.m.MessageToast.show("No file uploaded.");
            return; // Exit the function
          }
   
          var reader = new FileReader();
          var that = this; // Preserve the reference to the controller for use inside the FileReader's onload function
   
          reader.onload = function (e) {
            try {
              var data = new Uint8Array(e.target.result);
              var workbook = XLSX.read(data, { type: "array" });
              var worksheet = workbook.Sheets[workbook.SheetNames[0]];
              var jsonData = XLSX.utils.sheet_to_json(worksheet);
   
              // Create a JSON model and set the data
              var oModel = new sap.ui.model.json.JSONModel();
              oModel.setData(jsonData);
   
              // Set the model to the view
              var oView = that.getView();
              oView.setModel(oModel);
              console.log("Excelsheet Data ", jsonData);
   
              // Show success message
              sap.m.MessageToast.show(
                "Excel file uploaded and parsed successfully"
              );
            } catch (error) {
              // Log the error
              console.error("Error parsing Excel file:", error);
              // Show error message to the user
              sap.m.MessageToast.show(
                "Error parsing Excel file. Please upload a valid Excel file."
              );
            }
          };
   
          reader.readAsArrayBuffer(file);
        },
   
        _createColumnListItem: function () {
          return new sap.m.ColumnListItem({
            cells: [
              new sap.m.Text({ text: "{Value}" }),
              new sap.m.Text({ text: "{Description}" }),
             
            ],
          });
        },
   
        onPreviewDialogClose: function () {
          this._oPreviewDialog.close();
        },
       
        onPreviewPress: function (oEvent) {
          var fileUploader = this.getView().byId("fileUploader");
          var file = fileUploader.oFileUpload.files[0];
          // Check if a file is selected
          if (!file) {
            // Show an error message to the user using MessageToast
            sap.m.MessageToast.show("No file uploaded.");
            return; // Exit the function
          }
          var fileUploader = this.getView().byId("fileUploader");
   
          var reader = new FileReader();
          var that = this; // Preserve the reference to the controller for use inside the FileReader's onload function
   
          reader.onload = function (e) {
            try {
              var data = new Uint8Array(e.target.result);
              var workbook = XLSX.read(data, { type: "array" });
              var worksheet = workbook.Sheets[workbook.SheetNames[0]];
              var jsonData = XLSX.utils.sheet_to_json(worksheet);
   
              // Check if column headers match the expected headers
              var expectedHeaders = [
                "Value",
                "Description",
                
              ];
              var actualHeaders = Object.keys(jsonData[0]); // Get the actual headers from the first row of data
   
              // Compare the expected headers with the actual headers
              var invalidColumns = [];
              expectedHeaders.forEach(function (header) {
                if (actualHeaders.indexOf(header) === -1) {
                  invalidColumns.push(header);
                }
              });
   
              // If any invalid columns found, show error message
              if (invalidColumns.length > 0) {
                var errorMessage =
                  "The following columns are missing in the uploaded file: " +
                  invalidColumns.join(", ");
                sap.m.MessageToast.show(errorMessage);
                return; // Exit the function
              }
   
              // Proceed with opening the preview dialog
              that._openPreviewDialog(jsonData);
            } catch (error) {
              // Log and handle any errors
              console.error("Error processing Excel file:", error);
              sap.m.MessageToast.show(
                "Error processing Excel file. Please upload a valid Excel file."
              );
            }
          };
   
          reader.readAsArrayBuffer(file);
        },
   
        _openPreviewDialog: function (jsonData) {
          // Create a JSON model and set the data
          var oModel = new sap.ui.model.json.JSONModel();
          oModel.setData(jsonData);
   
          // Set the model to the view
          var oView = this.getView();
   
          // Instantiate the dialog fragment
          if (!this._oPreviewDialog) {
            this._oPreviewDialog = sap.ui.xmlfragment(
              oView.getId(),
              "nauticalfe.fragments.countryUpdSheetDialog",
              this
            );
            oView.addDependent(this._oPreviewDialog);
          }
   
          // Bind data to the table in the dialog
          var oTable = this._oPreviewDialog.getContent()[0];
          oTable.setModel(oModel);
          oTable.bindAggregation("items", "/", this._createColumnListItem());
   
          // Open dialog
          this._oPreviewDialog.open();
        },
   
        onDownloadPress: function () {
          // Create dummy data for the template (replace with your actual template data)
          var templateData = [
            [
              "Value",
              "Description",
             
            ],
            // Add more rows as needed
          ];
   
          // Create spreadsheet
          var oSpreadsheet = new sap.ui.export.Spreadsheet({
            workbook: {
              columns: [
                // Define columns for the spreadsheet
                { label: "Value", property: "Value" },
                { label: "Description", property: "Description" },
               
              ],
              rows: {
                // Bind data rows to the template data
                path: "/",
              },
            },
            dataSource: templateData,
          });
   
          // Download the spreadsheet
          oSpreadsheet.build();
        },
      });
    }
  );
   