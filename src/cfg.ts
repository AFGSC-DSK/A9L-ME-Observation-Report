import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "EventName",
                    title: "Event Name",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "Topic",
                    title: "Topic",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "ObservedBy",
                    title: "Observed By",
                    type: Helper.SPCfgFieldType.User,
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleOnly,
                } as Helper.IFieldInfoUser,
                {
                    name: "Observation",
                    title: "Observation",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "ObservationDate",
                    title: "Observation Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateTime,
                    displayFormat: SPTypes.DateFormat.DateTime,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoDate,
                {
                    name: "Classification",
                    title: "Classification",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "(S)-Secret",
                    required: true,
                    choices: [
                        "(S)-Secret", "(CUI)-Controlled Unclassified Information"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SubmittedRecommendedOPR",
                    title: "Submitted Recommended OPR",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "DOTMLPF",
                    title: "DOTMLPF",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "D-Doctrine",
                    required: true,
                    showInNewForm: true,
                    choices: [
                        "D-Doctrine", "O-Organization", "T-Training", "M-Material",
                        "L-Leadership", "P-Personnel", "F-Facilities"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Discussion",
                    title: "Discussion",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "Recommendations",
                    title: "Recommendations",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "Implications",
                    title: "Implications",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,
                {
                    name: "Keywords",
                    title: "Keywords",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true
                } as Helper.IFieldInfoText,                
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "EventName", "Topic", "ObservedBy", "Observation", "ObservationDate",
                        "Classification", "SubmittedRecommendedOPR", "DOTMLPF", "Discussion", "Recommendations",
                        "Implications", "Keywords",
                    ]
                }
            ]
        }
    ]
});

// Adds the solution to a classic page
Configuration["addToPage"] = (pageUrl: string) => {
    // Add a content editor webpart to the page
    Helper.addContentEditorWebPart(pageUrl, {
        contentLink: Strings.SolutionUrl,
        description: Strings.ProjectDescription,
        frameType: "None",
        title: Strings.ProjectName
    }).then(
        // Success
        () => {
            // Load
            console.log("[" + Strings.ProjectName + "] Successfully added the solution to the page.", pageUrl);
        },

        // Error
        ex => {
            // Load
            console.log("[" + Strings.ProjectName + "] Error adding the solution to the page.", ex);
        }
    );
}