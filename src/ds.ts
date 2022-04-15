import { Components, ContextInfo, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

// Item
export interface IItem extends Types.SP.ListItem {
    Title: string;
    EventName: string;
    Topic: string;
    ObservedBy: { Id: Number; Title: string };
    Observation: string;
    ObservationDate: string;
    Classification: string;
    SubmittedRecommendedOPR: string;
    DOTMLPF: string;
    Discussion: string;
    Recommendations: string;
    Implications: string;
    Keywords: string;
    Status: string;
}

// Configuration
export interface IConfiguration {
    adminGroupName?: string;
    membersGroupName?: string;
}

/**
 * Data Source
 */
export class DataSource {
    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters(): PromiseLike<Components.ICheckboxGroupItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field
            Web(Strings.SourceUrl).Lists(Strings.Lists.Main).Fields("Classification").execute((fld: Types.SP.FieldChoice) => {
                let items: Components.ICheckboxGroupItem[] = [];

                // Parse the choices
                for (let i = 0; i < fld.Choices.results.length; i++) {
                    // Add an item
                    items.push({
                        label: fld.Choices.results[i],
                        type: Components.CheckboxGroupTypes.Switch
                    });
                }

                // Set the filters and resolve the promise
                this._statusFilters = items;
                resolve(items);
            }, reject);
        });
    }

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            this.load().then(() => {
                // Load the config file settings
                this.loadConfiguration().then(() => {
                    // Load the user name
                    this.loadSecurityGroupUrls().then(() => {
                        // Determine if the user is an admin
                        this.GetAdminStatus().then(() => {
                            // Load the status filters
                            this.loadStatusFilters().then(() => {
                                // Resolve the request
                                resolve();
                            }, reject);
                        }, reject);
                    }, reject);
                }, reject);
            }, reject);
        });
    }

    // Loads the list data
    private static _items: IItem[] = null;
    static get Items(): IItem[] { return this._items; }
    static load(): PromiseLike<IItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            Web(Strings.SourceUrl).Lists(Strings.Lists.Main).Items().query({
                GetAllItems: true,
                Expand: ["AttachmentFiles", "ObservedBy"],
                OrderBy: ["Status"],
                Select: ["*", "ObservedBy/Id", "ObservedBy/Title"],
                Top: 5000
            }).execute(
                // Success
                items => {
                    // Set the items
                    this._items = items.results as any;

                    // Resolve the request
                    resolve(this._items);
                },
                // Error
                () => { reject(); }
            );
        });
    }

    // Event Registration Permissions
    private static _eventRegPerms: Types.SP.BasePermissions;
    static get EventRegPerms(): Types.SP.BasePermissions { return this._eventRegPerms; };

    // Check if user is an admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }

    // Set Admin status
    private static GetAdminStatus(): PromiseLike<void> {
        return new Promise((resolve) => {
            if (this._cfg.adminGroupName) {
                Web().SiteGroups().getByName(this._cfg.adminGroupName).Users().getById(ContextInfo.userId).execute(
                    () => { this._isAdmin = true; resolve(); },
                    () => { this._isAdmin = false; resolve(); }
                )
            }
            else {
                Web().AssociatedOwnerGroup().Users().getById(ContextInfo.userId).execute(
                    () => { this._isAdmin = true; resolve(); },
                    () => { this._isAdmin = false; resolve(); }
                )
            }
        });
    }

    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }
    static loadConfiguration(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current web
            Web().getFileByServerRelativeUrl(Strings.ObservationReportConfig).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
                    let cfg = null;
                    try { cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                    catch { cfg = {}; }

                    // Set the configuration
                    this._cfg = cfg;

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Set the configuration to nothing
                    this._cfg = {} as any;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }

    // Security Groups
    private static _managerId: number = null;
    static get ManagersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._managerId; };
    private static _memberId: number = null;
    static get MembersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._memberId; };
    static get ListUrl(): string { return ContextInfo.webServerRelativeUrl + "/Lists/Dashboard/AllItems.aspx" };
    static loadSecurityGroupUrls(): PromiseLike<void> {
        return new Promise((resolve) => {
            let web = Web();

            // Load the owner's group
            let ownersGroup = DataSource.Configuration.adminGroupName;
            (ownersGroup ? Web().SiteGroups().getByName(ownersGroup) : Web().AssociatedOwnerGroup()).execute(group => {
                // Set the id
                this._managerId = group.Id;
            });

            // Load the member's group
            let membersGroup = DataSource.Configuration.membersGroupName;
            (membersGroup ? Web().SiteGroups().getByName(membersGroup) : Web().AssociatedMemberGroup()).execute(group => {
                // Set the id
                this._memberId = group.Id;
            });

            // Wait for the requests to complete
            web.done(() => { resolve(); });
        });
    }
}