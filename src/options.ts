import { InstallationRequired, LoadingDialog, Modal } from "dattatable";
import { DataSource, IItem } from "./ds";
import { Components, Helper, Utility } from "gd-sprest-bs";
import { calendarPlus } from "gd-sprest-bs/build/icons/svgs/calendarPlus";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { sliders } from "gd-sprest-bs/build/icons/svgs/sliders";
import * as moment from "moment";
import { Configuration } from "./cfg";
export class Options {

    dropdownMenu(el: HTMLElement, item: IItem) {
        let emailButton = Components.Button({
            el: el,
            text: "Send Email",
            className: "dropdownMenu",
            onClick: () => {
                this.sendEmail(item)
            }
        });
    }

    // Setter for email content info for report
    private reportInfo(item: IItem) {
        let summaryView = `<h3><u>Observation Summary</u></h3>
                    <p><strong>Title: </strong>${item.Title}</p>
                    <p><strong>Event Name: </strong>${item.EventName}</p>
                    <p><strong>Topic: </strong>${item.Topic}</p>
                    <p><strong>Observed By: </strong>${item.ObservedBy.Title}</p>
                    <p><strong>Observation: </strong/>${item.Observation}</p>
                    <p><strong>Observation Date: </strong>${moment(item.ObservationDate).format("MM-DD-YYYY HH:mm")}</p>
                    <p><strong>Classification: </strong>${moment(item.Classification).format("MM-DD-YYYY HH:mm")}</p>
                    <p><strong>Submitted Recommended OPR: </strong>${item.SubmittedRecommendedOPR}</p>
                    <p><strong>DOTMLPF: </strong>${item.DOTMLPF}</p>
                    <p><strong>Discussion: </strong>${item.Discussion}</p>
                    <p><strong>Recommendations: </strong>${item.Recommendations}</p>
                    <p><strong>Implications: </strong>${item.Implications}</p>
                    <p><strong>Keywords: </strong>${item.Keywords}</p><br/>`;

        return summaryView;
    }

    private sendEmail(item: IItem) {
        // Set the modal header
        Modal.setHeader("Send Email");

        // Create the form
        let form = Components.Form({
            controls: [
                {
                    name: "EmailSubject",
                    label: "Email Subject",
                    required: true,
                    errorMessage: "A subject is required to send an email.",
                    type: Components.FormControlTypes.TextField,
                    value: item.Title
                },
                {
                    name: "EmailBody",
                    label: "Email Body",
                    required: true,
                    errorMessage: "Content is required to send an email.",
                    rows: 10,
                    type: Components.FormControlTypes.TextArea,
                    value: ""
                } as Components.IFormControlPropsTextField
            ]
        });

        // Set the modal body
        Modal.setBody(form.el);

        // Set the modal footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Send",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let values = form.getValues();

                            // Close the modal
                            Modal.hide();

                            // Show a loading dialog
                            LoadingDialog.setHeader("Sending Email");
                            LoadingDialog.setBody("This dialog will close after the email is sent.");
                            LoadingDialog.show();

                            let recipients = [];
                            for(let i = 0; i < item.EmailRecipients.results.length; i++) {
                                recipients.push(item.EmailRecipients.results[i].EMail);
                            }
                            
                            // Send the email
                            Utility().sendEmail({
                                To: recipients,
                                Body: values["EmailBody"].replace(/\n/g, "<br />") + this.reportInfo(item),
                                Subject: values["EmailSubject"]
                            }).execute(() => {
                                // Close the loading dialog
                                LoadingDialog.hide();
                            });
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        Modal.hide();
                    }
                }
            ]
        }).el);

        // Display the modal
        Modal.show();
    }
}