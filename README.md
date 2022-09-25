# Realm-Sheet-Bot
A project done in service for Realm Control Inc. Several individual Google Apps Script based projects used for internal company work automation: 

1. A Bill of Material(BOM) Creator

* The main branch functions as a repository for the BOM creator, which is a sidebar for the provided xml file that can be used to import item manifests from one spreadsheet another for additions to the BOM sheet.

* Other uses include the addition and subtraction of rows and columns to the BOM without adjusting formulas in the cells. It allows for text input as custom items to the BOM, which can also be added to the local or remote spreadsheets containing the item manifest

2. A Tool for creating Pull Schedules.

* A Pull Schedule constitutes the set of materials in a given project necessary, including placement, name, wiring connection dependent on the job's requirements.

* Using a pair of reference sheets for the original items and new items on the pull schedule allows this script to add the new items and cross out using formatting the old items and provide a timestamp for it to occur.

3. A tool for creating TRXIO based inventory supply requirements.

* TRXIO is a cloud based inventory management solution. Using TRXIO csv imports this script obtains data concerning customer project reservations on inventory quantity and using the properly formatted sheet creates a detailed list of items reserved for a project by the user interfacing with the spreadsheet and refers to the TRXIO data to insure available quantity and accuracy of inventory

4. msSQL Database API

* A Python API for interfacing with an SAP based msSQL database and Google Cloud Services to obtain Spreadsheet information. Allows SQL queries to adapt Spreadsheet data dependent on the google account associated with the Google Spreadsheet it has access to.
