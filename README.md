# ExcelExportButton
The ExcelExportButton is a custom SAPUI5 control that extends the [sap.m.Button](https://sapui5.hana.ondemand.com/#/api/sap.m.Button), that makes it fast and easy to export data
from any [sap.m.ListBase](https://sapui5.hana.ondemand.com/#/api/sap.m.ListBase) element into a compliant Microsoft Excel Spreadsheet.. It has the ability to perform a pre-defined
automatic download to Excel, or provide the user with a selection dialog that allows them to select which properties / rows to export.

The only requirement is that:
1. There is an element in the same view that extends [sap.m.ListBase](https://sapui5.hana.ondemand.com/#/api/sap.m.ListBase)
2. The ID of this element is defined in the ExcelExportButton's only required property 'listBaseID'. 
   If this element cannot be found in the view, or is not a [sap.m.ListBase](https://sapui5.hana.ondemand.com/#/api/sap.m.ListBase), the button will be hidden.
   
The most basic example to export all properties and rows on a single click is

```XML
<comp:ExcelExportButton listBaseID="theList"/>
```

The most basic example to give the user a full featured selection dialog is

```XML
<comp:ExcelExportButton listBaseID="theList" showSelectionDialog="true"/>
```

An example that genertes an automatic Excel with specific properties exported:

```XML
<comp:ExcelExportButton listBaseID="theList">
  <comp:properties>
    <comp:ExcelExportProperty label="First Name" value="firstName" type="sap.ui.export.EdmType.String"/>
    <comp:ExcelExportProperty label="Last Name" value="lastName" type="sap.ui.export.EdmType.String"/>
    <comp:ExcelExportProperty label="Date of Birth" value="dob" type="sap.ui.export.EdmType.String"/>
  </comp:properties>
</comp:ExcelExportButton>
```
    
Additional documentation and examples can be found in the [ExcelExportButton.pdf](https://github.com/tommycole6/ExcelExportButton/blob/main/Excel%20Export%20Button.pdf).
