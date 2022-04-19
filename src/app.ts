import { Dashboard, ItemForm, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { DataSource, IItem } from "./ds";
import Strings from "./strings";
import * as moment from "moment";
import { Options } from "./options"


/**
 * Main Application
 */
export class App {
    // Global Variables
    private _isAdmin: boolean = DataSource.IsAdmin;
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.Main;

        // Set the global variable
        this._el = el;

        // Render the dashboard
        this.render();
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the events
        DataSource.init().then(() => {
            // Clear the element
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render the dashboard
            this.render();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private render() {
        let options = new Options();

        // Create the dashboard
        let dashboard = new Dashboard({
            el: this._el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        dashboard.filter(2, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
                showFilter: false,
                items: this._isAdmin ? [
                    {
                        className: "btn-outline-light",
                        text: "Submit",
                        isButton: true,
                        onClick: () => {
                            // Create an item
                            ItemForm.create({
                                onUpdate: () => {
                                    // Load the data
                                    DataSource.load().then(items => {
                                        // Refresh the table
                                        dashboard.refresh(items);
                                    });
                                },
                                // onCreateEditForm: (props) => {
                                //     props.onControlRendering = (ctrl, field) => {
                                //         if (field.InternalName == "Title") {
                                //             field.InternalName = "Event Name";
                                //         }
                                //     }
                                //     return props;
                                // },
                                onSetFooter: (elFooter) => {
                                    // Render the close button
                                    Components.Button({
                                        el: elFooter,
                                        text: "Close",
                                        type: Components.ButtonTypes.OutlineSecondary,
                                        onClick: (button) => {
                                            ItemForm.close();
                                        }
                                    });
                                }
                            });
                        },
                    }
                ] : []
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.Items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 0,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        jQuery(api.context[0].nTable).removeClass('no-footer');
                        jQuery(api.context[0].nTable).addClass('tbl-footer');
                        jQuery(api.context[0].nTable).addClass('table-striped');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[2, "asc"]]
                },
                columns: [
                    {
                        name: "",
                        title: "Title",
                        onRenderCell: (el, column, item: IItem) => {
                            // Displays clickable title
                            let elViewLink = document.createElement("a");
                            elViewLink.href = "#";
                            el.appendChild(elViewLink);
                            let value = item.Title;
                            elViewLink.innerHTML = value;
                            elViewLink.className = "fw-bold text-primary";
                        },
                        onClickCell: (el, col, item: IItem) => {
                            // Show edit form if admin and view form if not
                            if (this._isAdmin) {
                                // Show the edit form
                                ItemForm.edit({
                                    itemId: item.Id,
                                    onUpdate: () => {
                                        // Refresh the data
                                        DataSource.load().then(items => {
                                            // Update the data
                                            dashboard.refresh(items);
                                        });
                                    },
                                    onSetFooter: (elFooter) => {
                                        // Render the close button
                                        Components.Button({
                                            el: elFooter,
                                            text: "Cancel",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: (button) => {
                                                ItemForm.close();
                                            }
                                        });
                                    }
                                });
                            } else {
                                // Show the edit form
                                ItemForm.view({
                                    itemId: item.Id,
                                    onSetFooter: (elFooter) => {
                                        // Render the close button
                                        Components.Button({
                                            el: elFooter,
                                            text: "Close",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: (button) => {
                                                ItemForm.close();
                                            }
                                        });
                                    }
                                });
                            }
                        }
                    },
                    {
                        name: "Observation ID",
                        title: "ObservationID",
                        onRenderCell: (el, column, item: IItem) => {
                            el.innerHTML = item.ObservationID ? item.ObservationID : "No Observation ID specified";
                        }
                    },
                    {
                        name: "Status",
                        title: "Status"
                    },
                    {
                        name: "EventName",
                        title: "Event Name",
                    },
                    {
                        name: "Observation",
                        title: "Observation",
                    },
                    {
                        name: "Discussion",
                        title: "Discussion"
                    },
                    {
                        name: "Recommendations",
                        title: "Recommendations"
                    },
                    {
                        name: "Modified By",
                        title: "",
                        onRenderCell: (el, column, item: IItem) => {
                            let modUser = item.Editor.Title;

                            el.innerHTML = modUser;
                        }
                    }
                    // {
                    //     name: "",
                    //     title: "Email",
                    //     onRenderCell: (el, column, item: IItem) => {
                    //         options.dropdownMenu(el, item);
                    //     }
                    // }
                ]
            }
        });
    }
}