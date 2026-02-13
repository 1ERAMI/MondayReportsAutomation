"""
Unified Monday Reports UI
Launch all Monday report automations from a single interface.
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar, BooleanVar, messagebox
import threading
import traceback

# Import report modules (standard imports, no more importlib needed)
import report_andy_greg as AndyGregReports
import report_cameron_flatirons as FlatironsReports
import report_cameron_crump as CrumpReports
import report_malissa as MalissaReports

# Import centralized config
from report_config import REPORTS, get_default_emails


class UnifiedReportSenderUI:
    """Unified UI for selecting and sending all Monday reports"""

    def __init__(self):
        # Build report configurations from centralized config
        self.report_configs = {
            REPORTS["andy_greg"]["display_name"]: {
                "module": AndyGregReports,
                "count": len(REPORTS["andy_greg"]["subject_filters"]),
                "emails": get_default_emails("andy_greg"),
            },
            REPORTS["cameron_flatirons"]["display_name"]: {
                "module": FlatironsReports,
                "count": len(REPORTS["cameron_flatirons"]["subject_filters"]),
                "emails": get_default_emails("cameron_flatirons"),
            },
            REPORTS["cameron_crump"]["display_name"]: {
                "module": CrumpReports,
                "count": len(REPORTS["cameron_crump"]["subject_filters"]),
                "emails": get_default_emails("cameron_crump"),
            },
            REPORTS["malissa"]["display_name"]: {
                "module": MalissaReports,
                "count": len(REPORTS["malissa"]["subject_filters"]),
                "emails": get_default_emails("malissa"),
            },
        }

        self.selected_report = None
        self.email_vars = {}
        self.report_vars = {}

        # Create main window
        self.root = ttk.Window(themename="darkly")
        self.root.title("Monday Reports - Universal Sender")
        self.root.geometry("700x850")

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Create StringVar after root window exists
        self.status_var = StringVar(value="Ready to send reports")
        self.send_email_var = BooleanVar(value=False)  # Email checkbox (default OFF)
        self.upload_drive_var = BooleanVar(value=True)  # Drive upload checkbox (default ON)

        self.setup_ui()

    def setup_ui(self):
        """Setup the UI components"""

        # Header
        header_frame = ttk.Frame(self.root, padding=20)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_label = ttk.Label(
            header_frame,
            text="\U0001f4ca Monday Reports - Universal Sender",
            font=("Segoe UI", 20, "bold"),
            bootstyle="inverse-primary"
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            header_frame,
            text="Select report type, choose recipients, and send",
            font=("Segoe UI", 10)
        )
        subtitle_label.pack(pady=(5, 0))

        # Main content area with scrollbar
        content_frame = ttk.Frame(self.root)
        content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # Canvas and scrollbar
        canvas = ttk.Canvas(content_frame)
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=canvas.yview, bootstyle="primary-round")
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Report Type Selection
        report_label = ttk.Label(
            scrollable_frame,
            text="\U0001f4cb Select Report Type(s):",
            font=("Segoe UI", 14, "bold"),
            bootstyle="info"
        )
        report_label.pack(anchor="w", padx=20, pady=(10, 5))

        # Add select/deselect all for reports
        report_ctrl_frame = ttk.Frame(scrollable_frame)
        report_ctrl_frame.pack(fill="x", padx=30, pady=2)

        select_all_reports_btn = ttk.Button(
            report_ctrl_frame,
            text="Select All Reports",
            command=self.select_all_reports,
            bootstyle="info-outline",
            width=18
        )
        select_all_reports_btn.pack(side="left", padx=(0, 5))

        deselect_all_reports_btn = ttk.Button(
            report_ctrl_frame,
            text="Deselect All Reports",
            command=self.deselect_all_reports,
            bootstyle="secondary-outline",
            width=18
        )
        deselect_all_reports_btn.pack(side="left")

        # Create checkboxes for each report type (allow multiple selection)
        for report_name, config in self.report_configs.items():
            var = BooleanVar(value=False)
            self.report_vars[report_name] = var

            cb_frame = ttk.Frame(scrollable_frame)
            cb_frame.pack(fill="x", padx=30, pady=5)

            cb = ttk.Checkbutton(
                cb_frame,
                text=f"{report_name} ({config['count']} reports)",
                variable=var,
                command=self.on_report_change,
                bootstyle="primary-round-toggle"
            )
            cb.pack(anchor="w")

        # Separator
        sep = ttk.Separator(scrollable_frame, bootstyle="secondary")
        sep.pack(fill="x", padx=20, pady=20)

        # Email Recipients Section
        self.email_section_label = ttk.Label(
            scrollable_frame,
            text="\U0001f4e7 Select Recipients:",
            font=("Segoe UI", 14, "bold"),
            bootstyle="info"
        )
        self.email_section_label.pack(anchor="w", padx=20, pady=(5, 10))

        # Frame to hold email checkboxes (will be rebuilt on report change)
        self.email_checkboxes_frame = ttk.Frame(scrollable_frame)
        self.email_checkboxes_frame.pack(fill="x", padx=30, pady=5)

        # Build initial email checkboxes
        self.build_email_checkboxes()

        # Pack canvas and scrollbar
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Delivery Options Section
        delivery_frame = ttk.LabelFrame(
            self.root,
            text="\U0001f4ec Delivery Options"
        )
        delivery_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(10, 0))

        # Inner frame for padding
        delivery_inner = ttk.Frame(delivery_frame)
        delivery_inner.pack(fill="both", expand=True, padx=20, pady=10)

        # Email checkbox
        email_checkbox = ttk.Checkbutton(
            delivery_inner,
            text="\U0001f4e7 Send via Email",
            variable=self.send_email_var,
            bootstyle="primary-round-toggle"
        )
        email_checkbox.pack(anchor="w", pady=2)

        # Drive checkbox
        drive_checkbox = ttk.Checkbutton(
            delivery_inner,
            text="\U0001f4e4 Upload to Google Drive",
            variable=self.upload_drive_var,
            bootstyle="success-round-toggle"
        )
        drive_checkbox.pack(anchor="w", pady=2)

        # Help text
        delivery_help = ttk.Label(
            delivery_inner,
            text="Select one or both delivery methods. Configure Drive folder IDs in drive_uploader.py",
            font=("Segoe UI", 8),
            bootstyle="secondary"
        )
        delivery_help.pack(anchor="w", pady=(5, 0))

        # Status bar (StringVar already created in __init__)
        self.status_label = ttk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bootstyle="inverse-secondary",
            padding=10
        )
        self.status_label.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))

        # Buttons frame
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=10)

        # Select/Deselect All buttons for recipients
        select_all_btn = ttk.Button(
            button_frame,
            text="\u2713 Select All Recipients",
            command=self.select_all,
            bootstyle="info-outline",
            width=20
        )
        select_all_btn.pack(side="left", padx=5)

        deselect_all_btn = ttk.Button(
            button_frame,
            text="\u2717 Deselect All Recipients",
            command=self.deselect_all,
            bootstyle="secondary-outline",
            width=20
        )
        deselect_all_btn.pack(side="left", padx=5)

        # Send button
        self.send_btn = ttk.Button(
            button_frame,
            text="\U0001f680 Send Reports",
            command=self.on_send,
            bootstyle="success",
            width=20
        )
        self.send_btn.pack(side="right", padx=5)

        # Cancel button
        cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.root.quit,
            bootstyle="danger-outline",
            width=15
        )
        cancel_btn.pack(side="right", padx=5)

    def build_email_checkboxes(self):
        """Build email checkboxes based on selected report(s)"""
        # Clear existing checkboxes
        for widget in self.email_checkboxes_frame.winfo_children():
            widget.destroy()

        self.email_vars.clear()

        # Get emails for all selected reports (union of all emails)
        selected_reports = [name for name, var in self.report_vars.items() if var.get()]

        if not selected_reports:
            # No reports selected, show message
            msg = ttk.Label(
                self.email_checkboxes_frame,
                text="Select at least one report type above",
                font=("Segoe UI", 9, "italic"),
                bootstyle="secondary"
            )
            msg.pack(anchor="w", pady=10)
            return

        # Collect all unique emails from selected reports
        all_emails = set()
        for report_name in selected_reports:
            all_emails.update(self.report_configs[report_name]["emails"])

        # Sort for consistent display
        emails = sorted(all_emails)

        # Create checkboxes
        for email in emails:
            var = BooleanVar(value=False)
            self.email_vars[email] = var

            cb = ttk.Checkbutton(
                self.email_checkboxes_frame,
                text=email,
                variable=var,
                bootstyle="primary-round-toggle"
            )
            cb.pack(anchor="w", pady=2)

    def on_report_change(self):
        """Handle report type selection change"""
        selected_reports = [name for name, var in self.report_vars.items() if var.get()]

        if not selected_reports:
            self.status_var.set("Select at least one report type")
        else:
            total_count = sum(self.report_configs[name]["count"] for name in selected_reports)
            self.status_var.set(f"Selected {len(selected_reports)} report type(s) - {total_count} total reports")

        self.build_email_checkboxes()

    def select_all_reports(self):
        """Select all report type checkboxes"""
        for var in self.report_vars.values():
            var.set(True)
        self.on_report_change()

    def deselect_all_reports(self):
        """Deselect all report type checkboxes"""
        for var in self.report_vars.values():
            var.set(False)
        self.on_report_change()

    def select_all(self):
        """Select all email recipient checkboxes"""
        for var in self.email_vars.values():
            var.set(True)
        self.status_var.set(f"Selected all {len(self.email_vars)} recipients")

    def deselect_all(self):
        """Deselect all email recipient checkboxes"""
        for var in self.email_vars.values():
            var.set(False)
        self.status_var.set("All recipients deselected")

    def on_send(self):
        """Handle send button click"""
        # Get selected reports
        selected_reports = [name for name, var in self.report_vars.items() if var.get()]

        if not selected_reports:
            messagebox.showwarning(
                "No Reports Selected",
                "Please select at least one report type."
            )
            return

        # Get selected emails
        selected_emails = [email for email, var in self.email_vars.items() if var.get()]

        # Check delivery options
        send_email = self.send_email_var.get()
        upload_drive = self.upload_drive_var.get()

        if not send_email and not upload_drive:
            messagebox.showwarning(
                "No Delivery Method",
                "Please select at least one delivery method (Email or Drive)."
            )
            return

        if send_email and not selected_emails:
            messagebox.showwarning(
                "No Recipients",
                "Please select at least one recipient when sending email."
            )
            return

        # Build delivery method text
        delivery_methods = []
        if send_email:
            delivery_methods.append("Email")
        if upload_drive:
            delivery_methods.append("Google Drive")
        delivery_text = " and ".join(delivery_methods)

        # Confirm send
        total_reports = sum(self.report_configs[name]["count"] for name in selected_reports)
        confirm = messagebox.askyesno(
            "Confirm Processing",
            f"Process {len(selected_reports)} report type(s) ({total_reports} total reports) to {len(selected_emails)} recipient(s)?\n\n"
            f"Delivery: {delivery_text}\n\n"
            f"Reports:\n" + "\n".join(f"  \u2022 {name}" for name in selected_reports) + "\n\n"
            f"Recipients:\n" + "\n".join(f"  \u2022 {email}" for email in selected_emails)
        )

        if confirm:
            self.send_btn.config(state="disabled")
            self.status_var.set(f"Processing {len(selected_reports)} report type(s)...")

            # Run in thread with all selected reports
            thread = threading.Thread(
                target=self.run_all_processes,
                args=(selected_reports, selected_emails, send_email, upload_drive)
            )
            thread.start()

    def run_all_processes(self, selected_reports, selected_emails, send_email=True, upload_to_drive=False):
        """Run multiple report processes sequentially in a separate thread"""
        try:
            for idx, report_name in enumerate(selected_reports, 1):
                status_msg = f"Processing report type {idx}/{len(selected_reports)}: {report_name}"
                print(status_msg)
                self.status_callback(status_msg)

                config = self.report_configs[report_name]
                module = config["module"]

                # Run this report's main function with delivery options
                module.main(
                    to_emails=selected_emails,
                    status_callback=self.status_callback,
                    send_email=send_email,
                    upload_to_drive=upload_to_drive
                )

                print(f"\u2713 Completed: {report_name}")

            # All reports complete
            success_msg = f"Successfully sent {len(selected_reports)} report type(s)!"
            self.root.after(0, lambda: messagebox.showinfo("Success", success_msg))
            self.root.after(100, self.root.quit)
        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.send_btn.config(state="normal"))
        finally:
            self.root.after(0, lambda: self.status_var.set("Process complete"))

    def status_callback(self, message):
        """Update status from background thread"""
        self.root.after(100, lambda: self.status_var.set(message))

    def run(self):
        """Start the UI"""
        self.root.mainloop()


def main():
    """Launch the unified report sender UI"""
    ui = UnifiedReportSenderUI()
    ui.run()


if __name__ == "__main__":
    main()
