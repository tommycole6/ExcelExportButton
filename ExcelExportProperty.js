sap.ui.define([
  "sap/m/ListItemBase"
], function(Base) {
  "use strict";
  /**
   * Constructor for a new ExcelExportProperty.
   *
   * @param {string} [sId] Id for the new control, generated automatically if no id is given
   * @param {object} [mSettings] Initial settings for the new control
   * 
   * @class
   * The <code>ExcelExportProperty</code> control provides a way for the developer to define which properties from the model data should be exported.<br/>
   * 
   * <b>NOTE:</b> this control makes extensive use of the <code>sap.ui.export.Spreadsheet</code> and <code>sap.ui.export.EdmType</code> for the actual Excel 
   * generation, and therefore requires SAPUI5 version 1.50.0 or higher to function.
   *
   * @extends sap.m.ListItemBase
   *
   * @author Tom Cole, Linx-AS, L.L.C.
   * @version 1.0.0
   *
   * @constructor
   * @public
   */
  var ExcelExportProperty = Base.extend("ExcelExportProperty", {
	  metadata: {
		  properties: {
			  /**
			   * The text label to use for the column in the generated Excel spreadsheet
			   */
			  label: { type: "string" },
			  /**
			   * The property of the object to export into the column in the generated Excel spreadsheet
			   */
			  value: { type : "string" },
			  /**
			   * The data type to use in the column in the generated Excel spreadsheet
			   */
			  type: { type: "sap.ui.export.EdmType", defaultValue: sap.ui.export.EdmType.String },
			  /**
			   * Indicates if the column in the generated Excel spreadsheet should wrap it's text
			   */
			  wrap: { type: "boolean", defaultValue: false }
		  }
	  },
	  constructor: function(sID, mSettings) {
		  sap.m.ListItemBase.apply(this, arguments);
	  }
  });
  return ExcelExportProperty;
});