sap.ui.define([
  "sap/m/Button",
  "sap/m/ButtonRenderer",
  "sap/ui/export/Spreadsheet",
  "sap/ui/export/EdmType"
], function(Button, Renderer, Spreadsheet, EdmType) {
  "use strict";
  /**
   * Constructor for a new ExcelExportButton.
   *
   * @param {string} [sId] Id for the new control, generated automatically if no id is given
   * @param {object} [mSettings] Initial settings for the new control
   *
   * @class
   * The <code>ExcelExportButton</code> control provides a easy way to provide more advanced Excel Export process for <code>sap.m.ListBase</code> controls. 
   * Developers can dictate which properties are to be exported and can also give the user the ability to select which properties are to be exported by giving them
   * a property selection dialog. The developer can also decide to export all the available data in the model of the associated <code>sap.m.ListBase</code>, or only the currently displayed
   * items. This control will automatically disable when the associated <code>sap.m.ListBase</code> element has no rows of data.<br/>
   * <b>NOTE:</b> this control makes extensive use of the <code>sap.ui.export.Spreadsheet</code> and <code>sap.ui.export.EdmType</code> objects for the actual Excel 
   * generation, and therefore requires SAPUI5 version 1.50.0 or higher to function.
   *
   * @extends sap.m.Button
   *
   * @author Tom Cole, Linx-AS, L.L.C.
   * @version 1.0.0
   *
   * @constructor
   * @public
   */
  var ExcelExportButton = Button.extend(basePackage + ".components.ExcelExportButton", {
	  metadata: {
		properties: {
			/**
			 * The id of the sap.m.ListBase control to use for data in the Excel file. If this value is empty <i>or</i>
			 * a control with the give id cannot be found within the view, the button will be hidden and the error logged
			 * to the console.
			 */
			listBaseID: { type: "string" },
			/**
			 * The desired file name for the generated Excel file.
			 */
			fileName: { type: "string", defaultValue: "Export.xlsx" },
			/**
			 * The desired sheet name for the generated Excel file.
			 */
			sheetName: { type: "string", defaultValue: "Export Data" },
			/**
			 * Indicates whether all rows available in the model should be exported (true)
			 * or if only the data loaded should be exported (false). 
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to true, this property will be ignored
			 * and the user will have the option to select from the dialog.
			 */
			exportAllRows: { type: "boolean", defaultValue: false },
			/**
			 * Indicates whether all properties should be exported. If set to false, only those properties
			 * defined in the <code>properties</code> aggregation will be exported.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to true, this property will be ignored.
			 */
			exportAllColumns: { type: "boolean", defaultValue: false },
			/**
			 * Indicates whether a dialog box should be presented to the user to allow them to select which
			 * properties they want exported in the Excel file. If this value set to false, then only properties in the 
			 * <code>properties</code> aggregation will be exported. If the <code>properties</code> aggregation is empty then
			 * all properties in the associated data objects will be exported.
			 */
			showSelectionDialog: { type: "boolean", defaultValue: false },
			/**
			 * Indicates if only the properties defined in the values aggregation should be available for user
			 * selection. If set to false, then the user will be given the opportunity to select from all available
			 * properties in the associated data elements, but those defined in the <code>properties</code> aggregation will be selected,
			 * the others will be unselected.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored.
			 * <b>NOTE:</b> If the <code>properties</code> aggregation is empty, this property is ignored.
			 */
			showOnlyDefinedProperties: { type: "boolean", defaultValue: true },
			/**
			 * Indicates if defined properties should be placed at the top.<br/>
			 * <b>NOTE:</b> if set to true and <code>sortProperties</code> is set to true, then the defined properties will be
			 * sorted and displayed on top, and the remaining undefined properties will be sorted and shown beneath.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored.
			 */
			showDefinedPropertiesFirst: { type: "boolean", defaultValue: false },
			/**
			 * Indicates if the defined properties should be initially selected when the dialog is displayed.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored. 
			 */
			selectDefinedProperties: { type: "boolean", defaultValue: true },
			/**
			 * Indicates if defined properties should be highlighted.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored. 
			 */
			highlightDefinedProperties: { type: "boolean", defaultValue: false },
			/**
			 * Indicates if the property names should be displayed below the property label.
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored.
			 */
			showPropertyNames : { type: "boolean", defaultValue: false },
			/**
			 * Indicated if selectable properties should be sorted alphabetically.<br/>
			 * <b>NOTE:</b> If <code>showSelectionDialog</code> is set to false, this property is ignored.
			 */
			sortProperties : { type: "boolean", defaultValue: true },
		},
		defaultAggregation : "values",
		aggregations: {
			/**
			 * Defines the list of properties to include in the Export. If <code>showSelectionDialog</code> is set to true and <code>showOnlyDefinedProperties</code>
			 * is set to false, the user will have the ability to select and export additional properties as defined in the data objects from the associated
			 * sap.ui.ListBase control's model.
			 */
			properties: { type : "ExcelExportProperty", multiple : true, singularName : "property" }
		}
   	  },
	  renderer: Renderer,
	  constructor: function(sID, mSettings) {
		  sap.m.Button.apply(this, arguments);
		  //assign default icon if one was not defined...
		  if (! this.getIcon()) {
			  this.setIcon("sap-icon://excel-attachment");
		  }
		  //assign default tooltip if one was not defined...
		  if (! this.getTooltip()) {
			  this.setTooltip("Download Excel");
		  }
		  //If no listBaseID property was provided display an error in the console and hide the button.
		  if (! this.getListBaseID() || this.getListBaseID().length < 1) {
			  console.error("ExcelExportButton: Property listBaseID cannot be empty.");
			  this.setVisible(false);
		  }
		  else {
			  var that = this;
			  setTimeout(function() {
				  try {
					  var oParent = that.getParent();
					  while (! (oParent instanceof sap.ui.core.mvc.View)) {
						  oParent = oParent.getParent();
					  }
					  that._oView = oParent;
					  let oList = that._oView.byId(that.getListBaseID());
					  //If no element with the provided listBaseID can be found in the parent view, display an error in the console and hide the button.
					  if (! oList) {
						  console.error("ExcelExportButton: Cannot find linked element with id '" + that.getListBaseID() + "'.");
						  that.setVisible(false);
					  }
					  else {
						  that._oList = oList;
						  that._checkEnabled();
						  that._oList.attachUpdateFinished(function(oEvent) {
							 that._checkEnabled(); 
						  });
					  }
				  }
				  catch(e) {
					  console.error(e);
				  }
			  }, 100);
			  this.attachPress(function(oEvent) {
				  that._processClick(oEvent);
			  });
		  }
	  },
	  _pDialog: null,
	  _oView: null,
	  _oList: null,
	  _aProps: [],
	  _checkEnabled: function() {
		  try {
			  this.setEnabled(this._oList.getItems().length > 0);
		  }
		  catch(e) {
			  
		  }
	  },
	  /**
	   * Event handler that is triggered when the ExcelExportButton is pressed.
	   * 
	   * If <code>showSelectionDialog</code> is set to true, it will generate and display the property selection dialog.
	   * The property selection dialog allows users to select which properties from the underlying data they want to have exported
	   * to Excel. There will be a Select All checkbox that allows the user to select and de-select all properties quickly.
	   * 
	   * If <code>showSelectionDialog</code> is set to false, this thie function forward to the <code>_generateExcel</code> function.
	   * 
	   * The contents of the dialog is based on the following settings: 
	   * 	<ul>
	   * 		<li>
	   * 			<code>showOnlyDefinedProperties</code>
	   * 			<blockquote>
	   * 				If set to true, only the properties defined in the <code>properties</code> aggregation will be displayed for selection.
	   * 				If set to false, a secondary checkbox is provided to allow the the user to select all Default properties. Default properties
	   * 				are defined as the properties listed in the <code>properties</code> aggregation.
	   * 				If the <code>properties</code> aggregation is empty, this property is ignored.
	   * 			</blockquote>
	   * 		</li>
	   * 		<li>
	   * 			<code>sortProperties</code>
	   * 			<blockquote>
	   * 				This will sort the properties in the selection dialog alphabetically by their label.
	   * 			</blockquote>
	   * 		</li>
	   * 		<li>
	   * 			<code>showDefinedPropertiesFirst</code>
	   * 			<blockquote>
	   * 				Properties defined in the <code>properties</code> aggregation will be displayed first.
	   * 			</blockquote>
	   * 		</li>
	   * 		<li>
	   * 			<code>selectDefinedProperties</code>
	   * 			<blockquote>
	   * 				Properties defined in the <code>properties</code> aggregation will be initially selected.
	   * 				If the <code>properties</code> aggregation is empty, this property is ignored.
	   * 			</blockquote>
	   * 		</li>
	   * 	</ul>
	   * @private
	   */
	  _processClick: function(oEvent) {
		  if (this._oList) {
			  if (this.getShowSelectionDialog()) {
				  //If this is the first time, generate the selection dialog element...
				  if (! this._pDialog) {
					  if (! this._oView) {
						  this._oView = this.getParent();
						  while (! (this._oView instanceof sap.ui.core.mvc.View)) {
							  this._oView = this._oView.getParent();
						  }
					  }
					  var aProperties = {
							  properties: []
					  };
					  if (this.getProperties().length > 0 && this.getShowOnlyDefinedProperties()) {
						  for (let i = 0; i < this.getProperties().length; i++) {
							  aProperties.properties.push({
								 title: this.getProperties()[i].getLabel(),
								 description: this.getProperties()[i].getValue(),
								 info: this.getProperties()[i].getType()
							  });
						  }
					  }
					  else {
						  let oList = this._oView.byId(this.getListBaseID());
						  let oModel = oList.getModel();
						  var oTemplate = oModel.getProperty(oList.getItems()[0].getBindingContext().getPath());
						  for (var key in oTemplate) {
							  let currentTopIndex = 0;
							  if (oTemplate.hasOwnProperty(key) && (! Array.isArray(oTemplate[key]))) {
								  let isDefined = false;
								  let definedLabel = this._generateUserFriendlyName(key);
								  if (this.getProperties().length > 0 && this.getShowDefinedPropertiesFirst()) {
									  for (let i = 0; i < this.getProperties().length; i++) {
										  if (this.getProperties()[i].getValue() === key) {
											  isDefined = true;
											  definedLabel = this.getProperties()[i].getLabel();
											  break;
										  }
									  }
								  }
								  aProperties.properties.push({
									 title: definedLabel,
									 description: key,
									 info: sap.ui.export.EdmType.String,
									 defined: isDefined
								  });
							  }
						  }
					  }
					  if (this.getSortProperties()) {
						  aProperties.properties.sort((a, b) => (a.title > b.title) ? 1 : -1);
					  }
					  if (this.getProperties().length > 0 && this.getShowDefinedPropertiesFirst() && ! this.getShowOnlyDefinedProperties()) {
						  aProperties.properties.sort(function(a, b) {
							  if (a.defined && ! b.defined) {
								  return -1;
							  }
							  else if (! a.defined && b.defined) {
								  return 1;
							  }
							  else {
								  return 0;
							  }
						  });
					  }
					  var that = this;
					  let pModel = new sap.ui.model.json.JSONModel(aProperties);
					  sap.ui.getCore().setModel(pModel, "properties");
					  var pTemplate = new sap.m.StandardListItem({
						  title: "{properties>title}",
					      description: this.getShowPropertyNames() ? "{properties>description}" : "",
					  });
					  var theListId = this._oView.createId("pList");
					  var theButtonId = this._oView.createId("selectAllPropertiesButton");
					  var theButtonId2 = this._oView.createId("selectDefinedPropertiesButton");
					  var theExportButtonId = this._oView.createId("exportAllButton");
					  var theExportButtonId2 = this._oView.createId("exportLoadedButton");
					  this._pDialog = new sap.m.Dialog("pDialog", {
						  title: "Select Values to Export",
						  stretch: sap.ui.Device.browser.mobile,
						  contentWidth: "600px",
						  subHeader: [
							  new sap.m.Toolbar({
								  content: [
									  new sap.m.CheckBox(theButtonId, {
										  text: "Select All (" + aProperties.properties.length + ")",
										  valueState: (this.getHighlightDefinedProperties()) ? "Information" : "None",
										  select: function(oEvent) {
											  that._toggleAllProperties(this);
										  }
									  }),
									  new sap.m.CheckBox(theButtonId2, {
										  text: "Select Defaults (" + that.getProperties().length + ")",
										  valueState: (this.getHighlightDefinedProperties()) ? "Success" : "None",
										  select: function(oEvent) {
											  that._toggleDefinedProperties(this);
										  }
									  })
								  ]
							  })
						  ],
						  content: [
							  new sap.m.List(theListId, {
								  noDataText: "No Properties Found",
								  rememberSelections: true,
								  mode: "MultiSelect",
								  includeItemInSelection: true
							  }).bindAggregation("items", "properties>/properties", pTemplate)
						  ],
						  buttons: [
							  new sap.m.Button(theExportButtonId, {
								  icon: "sap-icon://excel-attachment",
								  text: "Export All Rows",
								  press: function(oEvent) {
									  that._startExport(true);
								  }
							  }),
							  new sap.m.Button(theExportButtonId2, {
								  icon: "sap-icon://excel-attachment",
								  text: "Export Visible Rows",
								  press: function(oEvent) {
									  that._startExport(false);
								  }
							  }),
							  new sap.m.Button({
								  icon: "sap-icon://sys-cancel",
								  text: "Close",
								  press: function(oEvent) {
									  that._closePDialog();
								  }
							  })
						  ]
					  });
				  }
				  let sPath = "";
				  try {
					 sPath = this._oList.getBindingContext().getPath() + "/";
				  }
				  catch(e) {
					 
				  }
				  sPath += this._oList.getBindingInfo("items").path;
				  this._oView.byId("exportAllButton").setText("Export All Rows (" + this._oList.getModel().getProperty(sPath).length + ")");
				  this._oView.byId("exportLoadedButton").setText("Export Visible Rows (" + this._oList.getItems().length + ")");
				  this._pDialog.open();
				  //If there are defined properties and selectDefinedProperties was set to true, select them...
				  if (this.getProperties().length > 0 && this.getSelectDefinedProperties()) {
					  var that = this;
					  setTimeout(function() {
						  let theList = that._oView.byId("pList");
						  for (let i = 0; i < theList.getItems().length; i++) {
							  let theItem = theList.getItems()[i];
							  let theObject = theList.getModel("properties").getProperty(theItem.getBindingContextPath());
							  for (let j = 0; j < that.getProperties().length; j++) {
								  if (that.getProperties()[j].getValue() === theObject.description) {
									  theItem.setSelected(true);
									  if (that.getHighlightDefinedProperties()) {
										  theItem.setHighlight("Success");
										  theItem.setHighlightText("Default property");
									  }
									  break;
								  }
								  else if (that.getHighlightDefinedProperties()) {
									  theItem.setHighlight("Information");
									  theItem.setHighlightText("Discovered property");
								  }
							  }
						  }
						  that._oView.byId("selectDefinedPropertiesButton").setSelected(true);
					  }, 1);
				  }
				  //If no properties are in the properties aggregation OR showOnlyDefinedProperties is true, we can hide the select default checkbox...
				  if (this.getProperties().length < 1) {
					  this._oView.byId("selectDefinedPropertiesButton").setVisible(false);
				  } 
				  //If we are only showing defined properties, we can hide the select all checkbox...
				  else if (this.getShowOnlyDefinedProperties()) {
					  this._oView.byId("selectAllPropertiesButton").setVisible(false);
					  this._oView.byId("selectDefinedPropertiesButton").setText("Select All (" + this._oView.byId("pList").getItems().length + ")");
				  }
				  //if the table has the same amount of data as the model, remove the Export Defaults button...
				  if (this._oList.getModel().getProperty(sPath).length === this._oList.getItems().length) {
					  this._oView.byId("exportLoadedButton").setVisible(false);
				  }
			  }
			  else {
				  this._generateExcel();
			  }
		  }
	  },
	  /**
	   * Called by the Select All checkbox in the selection dialog. It will select or de-select all properties based on the checkbox status.
	   * 
	   * @private
	   */
	  _toggleAllProperties: function(oElement) {
		  if (! this._oView) {
			  this._oView = this.getParent();
			  while (! (this._oView instanceof sap.ui.core.mvc.View)) {
				  this._oView = this._oView.getParent();
			  }
		  }
		  var bSelected = oElement.getSelected();
		  var dSelected = this._oView.byId("selectDefinedPropertiesButton").getSelected();
		  let oList = this._oView.byId("pList");
		  if (oList) {
			    if (bSelected) {
			    	oList.selectAll();
				}
				else if (! dSelected){
					var items = oList.getSelectedItems();
					for (let i = 0; i < items.length; i++) {
						items[i].setSelected(false);
					}
				}
				else {
					var items = oList.getSelectedItems();
					for (let i = 0; i < items.length; i++) {
						var theProp = this.getShowPropertyNames() ? items[i].getDescription() : oList.getModel("properties").getProperty(items[i].getBindingContextPath()).description;
						var bFound = false;
						for (let j = 0; j < this.getProperties().length && ! bFound; j++) {
							if (theProp === this.getProperties()[j].getValue()) {
								items[i].setSelected(dSelected);
								bFound = true;
							}
						}
						if (! bFound) {
							items[i].setSelected(bSelected);
						}
					}
				}
		  }
	  },
	  /**
	   * Called by the Select Defaults checkbox in the selection dialog. It will select or de-select all properties 
	   * defined in the <code>properties</code> aggregation based on the checkbox status.
	   * 
	   * @private
	   */
	  _toggleDefinedProperties: function(oElement) {
		  if (! this._oView) {
			  this._oView = this.getParent();
			  while (! (this._oView instanceof sap.ui.core.mvc.View)) {
				  this._oView = this._oView.getParent();
			  }
		  }
		  var bSelected = oElement.getSelected();
		  let oList = this._oView.byId("pList");
		  if (oList) {
			  	var items = oList.getItems();
			  	for (let i = 0; i < items.length; i++) {
					var theProp = this.getShowPropertyNames() ? items[i].getDescription() : oList.getModel("properties").getProperty(items[i].getBindingContextPath()).description;
					for (let j = 0; j < this.getProperties().length; j++) {
						if (theProp === this.getProperties()[j].getValue()) {
							items[i].setSelected(bSelected);
							break;
						}
					}
			  	}
		  }
	  },
	  /**
	   * Called when the Export button in the property selection dialog is pressed.
	   * This generates a list of columns to export based on user selection and then calls
	   * <code>_generateExcel</code>.
	   * 
	   * @private
	   */
	  _startExport: function(bExportAll) {
		  if (! this._oView) {
			  this._oView = this.getParent();
			  while (! (this._oView instanceof sap.ui.core.mvc.View)) {
				  this._oView = this._oView.getParent();
			  }
		  }
		  let oList = this._oView.byId("pList");
		  if (oList) {
			  if (oList.getSelectedItems().length > 0) {
				  var aProperties = [];
				  for (let i = 0; i < oList.getSelectedItems().length; i++) {
					  let property = (this.getShowPropertyNames()) ? oList.getSelectedItems()[i].getDescription() : oList.getModel("properties").getProperty(oList.getSelectedItems()[i].getBindingContextPath()).description;
					  aProperties.push({
						  label: oList.getSelectedItems()[i].getTitle(),
						  property: property,
						  type: oList.getSelectedItems()[i].getInfo()
					  })
				  }
				  this._generateExcel(aProperties, bExportAll);
				  this._closePDialog();
			  }
			  else {
				  sap.m.MessageToast.show("Select at least one property to continue.");
			  }
		  }
	  },
	  /**
	   * Called by the Close button in the property selection dialog and closes it.
	   * 
	   * @private
	   */
	  _closePDialog: function() {
		  if (this._pDialog) {
			  this._pDialog.close();
		  }  
		  if (this._oView) {
			  this._oView.byId("selectAllPropertiesButton").setSelected(false);
			  if (this.getProperties().length > 0) {
				  this._oView.byId("selectDefinedPropertiesButton").setSelected(true);
			  }
		  }
	  },
	  /**
	   * Generates the row data from the linked <code>sap.ui.ListBase</code> element, gathering only the 
	   * properties defined. Then calls <code>_completeExcelExport</code> with the column and row data.
	   * 
	   * @private
	   */
	  _generateExcel: function(aProperties, bExportAll) {
		  if (! this._oView) {
			  this._oView = this.getParent();
			  while (! (this._oView instanceof sap.ui.core.mvc.View)) {
				  this._oView = this._oView.getParent();
			  }
		  }
		  var aProps = [];
		  //If aProperties is populated, then we have been called by the selection dialog and we can safely use those
		  //selected properties for our column list...
		  if (aProperties && aProperties.length > 0) {
			  aProps = aProperties;
		  }
		  //We have been called without a selection dialog. If we have properties defined and exportAllColumns is false,
		  //the use those defined properties as our column list...
		  else if (this.getProperties().length > 0 && ! this.getExportAllColumns()) {
			  for (let i = 0; i < this.getProperties().length; i++) {
				  aProps.push({
					  label: this.getProperties()[i].getLabel(),
					  property: this.getProperties()[i].getValue(),
					  type: this.getProperties()[i].getType(),
					  wrap: this.getProperties()[i].getWrap()
				  })
			  }
		  }
		  //We have been called without a selection dialog and we are exporting all properties as none have been selected
		  //or defined...
		  else {
			  let oList = this._oView.byId(this.getListBaseID());
			  let oModel = oList.getModel();
			  //Use the data attached to the first item in the ListBase as our data template to retrieve properties from...
			  var oTemplate = oModel.getProperty(oList.getItems()[0].getBindingContext().getPath());
			  for (var key in oTemplate) {
				  //Get all native properties from the template element...
				  if (oTemplate.hasOwnProperty(key) && (! Array.isArray(oTemplate[key]))) {
					  var bDefined = false;
					  //See if this property has a matching element in the properties aggregation.
					  //If yes, then use that label and type for export...
					  for (let i = 0; i < this.getProperties().length; i++) {
						  if (this.getProperties()[i].getValue() === key) {
							  bDefined = true;
							  aProps.push({
								  label: this.getProperties()[i].getLabel(),
								  property: key,
								  type: this.getProperties()[i].getType(),
								  wrap: this.getProperties()[i].getWrap()
							  });
							  break;
						  }
					  }
					  //Property does not have a matching property aggregation. Create a human readable label
					  //from the property name and set the type to the default String...
					  if (! bDefined) {
						  aProps.push({
							 label: this._generateUserFriendlyName(key),
							 property: key,
							 type: sap.ui.export.EdmType.String,
							 wrap: false
						  });
					  }
				  }
			  }
		  }
		  if (aProps.length > 0) {
			 //Pull the data from the model attached to the associated ListBase item for the properties found...
			 let oList = this._oView.byId(this.getListBaseID());
			 let aData = [];
			 if (bExportAll || (this.getExportAllRows() && (! this.getShowSelectionDialog()))) {
				 let sPath = "";
				 try {
					 sPath = oList.getBindingContext().getPath() + "/";
				 }
				 catch(e) {
					 
				 }
				 sPath += oList.getBindingInfo("items").path;
				 aData = oList.getModel().getProperty(sPath);
			 }
			 else {
				 let list = oList.getItems();
				 for (let i = 0; i < list.length; i++) {
					let item = list[i];
					let obj = item.getBindingContext().getProperty();
					aData.push(obj);
				 }
			 }
			 this._completeExcelExport(aData, aProps);
		  }
	  },
	  /**
	   * Used for properties that are in the selection list, but to not have a matching entry in the <code>properties</code> aggregation.
	   * It attempts to make a human readable label from a property name by:
	   * <ol>
	   * 	<li>Making the first character uppercase</li>
	   * 	<li>Replacing '_x' or '_X' with X</li>
	   * 	<li>Replace all 'xX' with ' X'</li>
	   * </ol>
	   * 
	   * Example:
	   * 	this_is_a_test 	-> This Is A Test
	   *	thisIsATest 	-> This Is A Test
	   *	this_isATest	-> This Is A Test
	   *    CheckUPSDate	-> Check UPS Date
	   *	check_UPS_date	-> Check UPS Date
	   * 
	   * @private
	   */	  
	  _generateUserFriendlyName(sLabel) {
		  let bits = sLabel.split("_");
		  let sNewLabel = "";
		  for (let i = 0; i < bits.length; i++) {
			  let bit = bits[i].trim();
			  sNewLabel += (bit.charAt(0).toUpperCase() + bit.slice(1)).trim();
		  }
		  return (sNewLabel.charAt(0).toUpperCase() + sNewLabel.slice(1)).replace(/([A-Z]+)/g, " $1").replace(/([A-Z][a-z])/g, " $1").replace(/  +/g, ' ').trim();
	  },
	  /**
	   * Takes the selected or define properties, the extracted row data and then uses the <code>sap.ui.export.Spreadsheet</code>
	   * object to generate the Excel spreadsheet and initate download.
	   * 
	   * @private
	   */
	  _completeExcelExport: function(aData, aProperties) {
			if (aData && aData.length > 0) {
				//Create a settings object for the sap.ui.exportSpreadsheet...
				var oSettings = {
						workbook: { 
							columns: aProperties,
							context: {
								sheetName: this.getSheetName()
							}
						},
						dataSource: aData,
						fileName: this.getFileName()
				};
				var oSheet = new sap.ui.export.Spreadsheet(oSettings);
				var that = this;
				//Generate spreadsheet, initiate download and close the dialog if open...
				oSheet.build().then(function() {
					that._closePDialog();
					MessageToast.show("Excel export finished");
				}).finally(function() {
					oSheet.destroy()
				});
			}
			else {
				sap.m.MessageToast.show("No data to export");
			}
	  }
  })
  return ExcelExportButton;
});