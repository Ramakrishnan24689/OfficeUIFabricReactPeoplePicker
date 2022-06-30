/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    peoplePicker: ComponentFramework.PropertyTypes.StringProperty;
    entityName: ComponentFramework.PropertyTypes.StringProperty;
    searchType: ComponentFramework.PropertyTypes.StringProperty;
    textFilterLength: ComponentFramework.PropertyTypes.WholeNumberProperty;
    includeDisabled: ComponentFramework.PropertyTypes.WholeNumberProperty;
    filterQuery: ComponentFramework.PropertyTypes.StringProperty;
    queryPrimaryField: ComponentFramework.PropertyTypes.StringProperty;
    querySecondaryField: ComponentFramework.PropertyTypes.StringProperty;
    queryResultIdField: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    peoplePicker?: string;
}
