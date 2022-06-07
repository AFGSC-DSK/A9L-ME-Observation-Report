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
                    name: "ObservationID",
                    title: "Observation ID",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoText,
                {
                    name: "EventName",
                    title: "Event Name",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                } as Helper.IFieldInfoText,
                {
                    name: "Topic",
                    title: "Topic",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoText,
                {
                    name: "ObservedBy",
                    title: "Observed By",
                    type: Helper.SPCfgFieldType.User,
                    required: false,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleOnly,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoUser,
                {
                    name: "Observation",
                    title: "Observation",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.EnhancedRichText,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                }  as Helper.IFieldInfoNote,
                {
                    name: "ObservationDate",
                    title: "Observation Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateTime,
                    displayFormat: SPTypes.DateFormat.DateTime,
                    defaultValue: "",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoDate,
                {
                    name: "Classification",
                    title: "Classification",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "(S)-Secret//No FORON",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                    choices: [
                        "(S)-Secret//No FORON", "(CUI)-Controlled Unclassified Information"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SubmittedRecommendedOPR",
                    title: "OPR",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                } as Helper.IFieldInfoText,
                {
                    name: "DOTMLPF",
                    title: "DOTMLPF",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "D-Doctrine",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                    choices: [
                        "D-Doctrine", "O-Organization", "T-Training", "M-Material",
                        "L-Leadership", "P-Personnel", "F-Facilities"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Discussion",
                    title: "Discussion",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.EnhancedRichText,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                } as Helper.IFieldInfoNote,
                {
                    name: "Recommendations",
                    title: "Recommendations",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.EnhancedRichText,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                } as Helper.IFieldInfoNote,
                {
                    name: "Implications",
                    title: "Implications",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoText,
                {
                    name: "Keywords",
                    title: "Keywords",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: false,
                    showInViewForms: false,
                    showInEditForm: false,
                    showInNewForm: false,
                } as Helper.IFieldInfoText, 
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "New",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                    choices: [
                        "New", "Valid", "In-valid", "In-progress", "Closed"
                    ]
                },
                {
                    name: "Comments",
                    title: "Comments",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.EnhancedRichText,
                    defaultValue: "",
                    required: false,
                    showInViewForms: true,
                    showInEditForm: true,
                    showInNewForm: true,
                } as Helper.IFieldInfoNote,            
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "ObservationID", "EventName", "Topic", "ObservedBy", "Observation", "ObservationDate",
                        "Classification", "SubmittedRecommendedOPR", "DOTMLPF", "Discussion", "Recommendations",
                        "Implications", "Keywords", "Status", "Comments"
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