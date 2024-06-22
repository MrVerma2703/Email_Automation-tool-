import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import asyncio
import threading
import random
import time

class EmailAutomationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation Tool")

        self.file_path = None
        self.sheets_info = {}
        self.event_loops = {}

        self.create_gui()

    def create_gui(self):
        # Select File Button
        select_file_btn = ttk.Button(self.root, text="Select Spreadsheet", command=self.select_file)
        select_file_btn.pack(pady=10)
        # Create a Canvas with a Vertical Scrollbar
        canvas_frame = ttk.Frame(self.root)
        canvas_frame.pack(expand=True, fill="both")
        self.canvas = tk.Canvas(canvas_frame, bg="white")
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # Create a Frame inside the Canvas to hold the controls
        self.sheets_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.sheets_frame, anchor="nw")

        # Bind the Canvas to the mouse wheel for scrolling
        self.canvas.bind_all("<MouseWheel>", lambda event: self.canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.file_path:
            self.setup_sheets_info()
            # Schedule the display_sheets function to be executed in the main thread
            self.root.after(0, self.display_sheets)

    def setup_sheets_info(self):
        # Read Excel file
        excel_data = pd.ExcelFile(self.file_path)
        sheet_names = excel_data.sheet_names

        # Initialize sheets_info dictionary
        self.sheets_info = {sheet_name: {"templates": [], "selected_template": tk.StringVar(), "template_combobox": None, "email_queue": [], "send_emails_btn": None} for sheet_name in sheet_names}

    def display_sheets(self):
        # Create controls for each sheet
        for sheet_name, info in self.sheets_info.items():
            sheet_frame = ttk.Frame(self.sheets_frame)
            sheet_frame.pack(pady=10, anchor=tk.W)

            # Sheet name label
            ttk.Label(sheet_frame, text=sheet_name).pack(side=tk.LEFT, padx=5)

            # Import Template Button
            import_template_btn = ttk.Button(sheet_frame, text="Import Template", command=lambda name=sheet_name: self.import_template(name))
            import_template_btn.pack(side=tk.LEFT, padx=5)

            # Select Templates
            ttk.Label(sheet_frame, text="Select Template:").pack(side=tk.LEFT, padx=5)
            template_combobox = ttk.Combobox(sheet_frame, textvariable=info["selected_template"], state="readonly")
            template_combobox.pack(side=tk.LEFT, padx=5)
            info["template_combobox"] = template_combobox

            # Remove Template Button
            remove_template_btn = ttk.Button(sheet_frame, text="Remove Template", command=lambda name=sheet_name: self.remove_template(name, template_combobox))
            remove_template_btn.pack(side=tk.LEFT, padx=5)

            # Send Emails Button
            send_emails_btn = ttk.Button(sheet_frame, text="Send Emails", command=lambda name=sheet_name: self.send_emails(name, template_combobox))
            send_emails_btn.pack(side=tk.LEFT, padx=5)
            info["send_emails_btn"] = send_emails_btn

        # Update the scroll region of the Canvas
        self.sheets_frame.update_idletasks()
        self.sheets_frame_width = self.sheets_frame.winfo_reqwidth()
        self.sheets_frame_height = self.sheets_frame.winfo_reqheight()
        self.sheets_frame_width += 10
        self.sheets_frame_height += 10
        self.canvas.config(scrollregion=(0, 0, self.sheets_frame_width, self.sheets_frame_height))

    def import_template(self, sheet_name):
        template_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if template_path:
            with open(template_path, "r") as file:
                template_content = file.read()
                self.sheets_info[sheet_name]["templates"].append({"name": os.path.basename(template_path), "content": template_content})
                # Schedule the update_template_combobox function to be executed in the main thread
                self.root.after(0, lambda: self.update_template_combobox(sheet_name))

    def remove_template(self, sheet_name, template_combobox):
        templates = self.sheets_info[sheet_name]["templates"]
        if templates:
            selected_template_index = template_combobox.current()
            if selected_template_index is not None:
                del self.sheets_info[sheet_name]["templates"][selected_template_index]
                self.sheets_info[sheet_name]["selected_template"].set("")  # Reset selected template
                # Schedule the update_template_combobox function to be executed in the main thread
                self.root.after(0, lambda: self.update_template_combobox(sheet_name))

    def update_template_combobox(self, sheet_name):
        templates = self.sheets_info[sheet_name]["templates"]
        template_combobox = self.sheets_info[sheet_name]["template_combobox"]
        template_names = [template["name"] for template in templates]
        template_combobox["values"] = template_names

    def send_emails(self, sheet_name, template_combobox):
        # Get the selected template name
        selected_template_name = self.sheets_info[sheet_name]["selected_template"].get()

        # Check if a template is selected
        if not selected_template_name:
            messagebox.showerror("Error", f"Please select a template for {sheet_name}.")
            return

        # Make the "Send Emails" button unresponsive during scheduling
        send_emails_btn = self.sheets_info[sheet_name]["send_emails_btn"]
        send_emails_btn.config(state=tk.DISABLED)

        # Start a new thread to run the asynchronous task
        thread = threading.Thread(target=self.run_async_task, args=(sheet_name, selected_template_name))
        thread.start()

    def run_async_task(self, sheet_name, selected_template_name):
        # Check if an event loop is already running for the sheet
        if sheet_name in self.event_loops:
            return

        # Create a new event loop for the thread
        self.event_loops[sheet_name] = asyncio.new_event_loop()

        try:
            # Run the asynchronous task
            asyncio.set_event_loop(self.event_loops[sheet_name])
            self.event_loops[sheet_name].run_until_complete(self.send_emails_async(sheet_name, selected_template_name))
        finally:
            # Close the event loop
            self.event_loops[sheet_name].close()
            asyncio.set_event_loop(None)

        # Remove the event loop after completing the task
        del self.event_loops[sheet_name]

    async def send_emails_async(self, sheet_name, selected_template_name):
        info = self.sheets_info[sheet_name]

        selected_template = next((t for t in info["templates"] if t["name"] == selected_template_name), None)

        if selected_template is None:
            return

        # Read sheet data
        df = pd.read_excel(self.file_path, sheet_name)

        sender_email = f"{sheet_name}@gmail.com"
        sender_password = df["Password"].iloc[0]  # Use the password from the first row of the sheet

        # Get the sender name from the D2 cell
        sender_name = str(df.loc[0, "Name"])
        # Check if sender_name is NaN, replace it with an empty string
        sender_name = sender_name if pd.notna(sender_name) else ""

        # Disable the "Send Emails" button
        #info["send_emails_btn"].config(state=tk.DISABLED)

        # Process emails in chunks of 50
        for chunk_start in range(0, len(df), 50):
            chunk_end = chunk_start + 50
            chunk_df = df.iloc[chunk_start:chunk_end]

            # Create a list to store asyncio tasks for each email
            tasks = []

            # Process the current chunk
            for index, row in chunk_df.iterrows():
                receiver_email = row["Emails"]
                receiver_name = self.extract_receiver_name(row["Websites Url"])     

                # Compose email
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = selected_template["content"].split("\n", 1)[0].replace("Sub", receiver_name)

                # Attach HTML body
                html_body = selected_template["content"].split("\n", 1)[1].replace("{sender_name}", sender_name)
                body = MIMEText(html_body, "html")
                message.attach(body)

                info["email_queue"].append({"sender_email": sender_email, "sender_password": sender_password, "receiver_email": receiver_email, "message": message})
                
                # Create a task for sending each email
                task = asyncio.create_task(self.send_email(sheet_name, sender_email, sender_password, receiver_email, message))

                # Introduce a fixed interval of 60 seconds between each email
                await asyncio.sleep(60)

                # Wait for the task to complete
                await task

            # Run all tasks concurrently
            await asyncio.gather(*tasks)

        # Enable the "Send Emails" button after completion
        info["send_emails_btn"].config(state=tk.NORMAL)

    async def send_email(self, sheet_name, sender_email, sender_password, receiver_email, message):
        try:
            # Send email
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, receiver_email, message.as_string())
        except smtplib.SMTPAuthenticationError:
            print(f"Authentication failed for {sender_email}. Please check the credentials.")

    def extract_receiver_name(self, url):
        # Extract receiver name from URL logic
        # Replace this with your appropriate logic
        return url.split(".")[1].capitalize()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAutomationTool(root)
    root.mainloop()
