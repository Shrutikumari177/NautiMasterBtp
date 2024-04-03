sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/Fragment",
    "sap/ui/core/routing/History",
    "sap/m/MessageToast",
    "sap/ui/export/Spreadsheet",
    "sap/ui/export/library"
  ],
  /**
   * @param {typeof sap.ui.core.mvc.Controller} Controller
   */
  function (Controller, Fragment, History, MessageToast, Spreadsheet) {
    "use strict";
 
    return Controller.extend("nauticalfe.controller.PortLocUpload", {
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
            new sap.m.Text({ text: "{Country}" }),
            new sap.m.Text({ text: "{Portc}" }),
            new sap.m.Text({ text: "{Portn}" }),
            new sap.m.Text({ text: "{Reancho}" }),
            new sap.m.Text({ text: "{Latitude}" }),
            new sap.m.Text({ text: "{Longitude}" }),
            new sap.m.Text({ text: "{Countryn}" }),
            new sap.m.Text({ text: "{LocID}" }),
            new sap.m.Text({ text: "{IND}" }),
          ],
        });
      },
 
      onPreviewDialogClose: function () {
        this._oPreviewDialog.close();
      },
     
      onPreviewPress1: function (oEvent) {
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
              "Country",
              "Portc",
              "Portn",
              "Reancho",
              "Latitude",
              "Longitude",
              "Countryn",
              "LocID",
              "IND",
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
            let errorIndex= [];
            for(let i = 0; i < jsonData.length; i++){
              console.log(typeof(jsonData[i].Latitude)==="number");
              if(typeof jsonData[i].Latitude === "number" && typeof jsonData[i].Longitude === "number"){
                continue;
              }
              else{
                errorIndex.push(i);
              }
            }
            if(errorIndex.length !== 0){
              let index = errorIndex.length === 1 ? "index" : "indices";
              let errorMessage = "";
              for(let i = 0; i < errorIndex.length - 1; i++){
                errorMessage = errorMessage + errorIndex[i] + ", "
              }
              errorMessage = errorMessage + errorIndex[errorIndex.length - 1];
              sap.m.MessageToast.show(`Latitude and Longitude value for Row ${index} : ${errorMessage} not as required`);
              return;
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
      onPreviewPress: function (oEvent) {
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
    
                // Check if column headers match the expected headers
                var expectedHeaders = [
                    "Country",
                    "Portc",
                    "Portn",
                    "Reancho",
                    "Latitude",
                    "Longitude",
                    "Countryn",
                    "LocID",
                    "IND",
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
    
                // Capitalize only the first letter of each field data
                jsonData.forEach(function (rowData) {
                    Object.keys(rowData).forEach(function (field) {
                        if (typeof rowData[field] === "string") {
                            rowData[field] = rowData[field].charAt(0).toUpperCase() + rowData[field].slice(1).toLowerCase();
                        }
                    });
                });
    
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
            "nauticalfe.fragments.previewSheetDialog",
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
        let sFileName =  "PortLocUpload"
        // Create dummy data for the template (replace with your actual template data)
        var templateData = [
          [
            "Country",
            "Portc",
            "Portn",
            "Reancho",
            "Latitude",
            "Longitude",
            "Countryn",
            "LocID",
            "IND",
          ],
          // Add more rows as needed
        ];
 
        // Create spreadsheet
        var oSpreadsheet = new sap.ui.export.Spreadsheet({
          workbook: {
            columns: [
              // Define columns for the spreadsheet
              { label: "Country", property: "Country" },
              { label: "Portc", property: "Portc" },
              { label: "Portn", property: "Portn" },
              { label: "Reancho", property: "Reancho" },
              { label: "Latitude", property: "Latitude" },
              { label: "Longitude", property: "Longitude" },
              { label: "Countryn", property: "Countryn" },
              { label: "LocID", property: "LocID" },
              { label: "IND", property: "IND" },
            ],
            rows: {
              // Bind data rows to the template data
              path: "/",
            },
          },
          dataSource: templateData,
          fileName: sFileName 
        });
 
        // Download the spreadsheet
        oSpreadsheet.build();
      },
    });
  }
);
 