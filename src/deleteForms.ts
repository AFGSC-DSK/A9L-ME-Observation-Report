import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, Utility } from "gd-sprest-bs";
import { IItem } from "./ds";

export class deleteForms {
    // Deletes an event
    static delete(eventItem: IItem, onRefresh: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Delete Event");

        // Set the body
        Modal.setBody("Are you sure you wanted to delete the selected event?");

        // Set the type
        Modal.setType(Components.ModalTypes.Medium);

        // Set the footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Yes",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Show the loading dialog
                        LoadingDialog.setHeader("Delete Event");
                        LoadingDialog.setBody("Deleting the event");
                        LoadingDialog.show();

                        // Delete the item
                        eventItem.delete().execute(
                            () => {
                                // Refresh the dashboard
                                onRefresh();

                                // Hide the dialog/modal
                                Modal.hide();
                                LoadingDialog.hide();
                            },
                            () => {
                                LoadingDialog.hide();
                            }
                        );
                    },
                },
                {
                    text: "No",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        // Hide the modal
                        Modal.hide();
                    },
                },
            ],
        }).el);

        // Show the modal
        Modal.show();
    }
}