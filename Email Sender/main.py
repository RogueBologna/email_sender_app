import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import json
import os
import traceback

class EmailSenderApp:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("Modern Email Sender")
            self.root.geometry("700x550")  # Smaller initial size for 3440x1440
            self.root.minsize(500, 400)   # Minimum size
            self.root.resizable(True, True)  # Allow resizing
            
            # Set window icon
            try:
                icon_path = os.path.join(os.path.dirname(__file__), "email_icon.ico")
                self.root.iconbitmap(icon_path)
            except Exception as e:
                with open("error.log", "a") as f:
                    f.write(f"Error setting icon: {str(e)}\n")
                    f.write(traceback.format_exc())
                # Continue without icon if it fails
            
            # Initialize state
            self.dark_mode = False
            self.credentials_file = "credentials.json"
            self.history_file = "email_history.json"
            self.templates = self.load_templates()
            self.attachment_path = None
            self.credentials = self.load_credentials()
            
            # Configure style for modern UI
            self.style = ttk.Style()
            self.style.theme_use('clam')
            
            # Main frame
            self.main_frame = ttk.Frame(self.root, padding="15 15 15 15", style="Main.TFrame")
            self.main_frame.grid(row=0, column=0, sticky="nsew")
            
            # Configure root grid weights for responsiveness
            self.root.grid_rowconfigure(0, weight=1)
            self.root.grid_columnconfigure(0, weight=1)
            
            # Configure main frame grid weights
            for i in range(10):
                self.main_frame.grid_rowconfigure(i, weight=1)
            self.main_frame.grid_columnconfigure(1, weight=1)
            
            # Title
            self.title_label = ttk.Label(self.main_frame, text="Modern Email Sender", font=("Segoe UI", 14, "bold"))
            self.title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky="ew")
            
            # Sender email and password
            self.sender_label = ttk.Label(self.main_frame, text="Your Email:")
            self.sender_label.grid(row=1, column=0, sticky="w", pady=5)
            self.sender_entry = ttk.Entry(self.main_frame, width=35)
            self.sender_entry.grid(row=1, column=1, columnspan=2, pady=5, sticky="ew")
            if self.credentials.get("email"):
                self.sender_entry.insert(0, self.credentials["email"])
            
            self.password_label = ttk.Label(self.main_frame, text="App Password:")
            self.password_label.grid(row=2, column=0, sticky="w", pady=5)
            self.password_entry = ttk.Entry(self.main_frame, width=35, show="*")
            self.password_entry.grid(row=2, column=1, columnspan=2, pady=5, sticky="ew")
            if self.credentials.get("password"):
                self.password_entry.insert(0, self.credentials["password"])
            
            self.save_credentials_var = tk.BooleanVar(value=bool(self.credentials))
            self.save_credentials_check = ttk.Checkbutton(
                self.main_frame, text="Save Email & Password", variable=self.save_credentials_var
            )
            self.save_credentials_check.grid(row=3, column=1, sticky="w", pady=5)
            
            # Recipient and subject
            self.recipient_label = ttk.Label(self.main_frame, text="Recipient Emails (comma-separated):")
            self.recipient_label.grid(row=4, column=0, sticky="w", pady=5)
            self.recipient_entry = ttk.Entry(self.main_frame, width=35)
            self.recipient_entry.grid(row=4, column=1, columnspan=2, pady=5, sticky="ew")
            
            self.subject_label = ttk.Label(self.main_frame, text="Subject:")
            self.subject_label.grid(row=5, column=0, sticky="w", pady=5)
            self.subject_entry = ttk.Entry(self.main_frame, width=35)
            self.subject_entry.grid(row=5, column=1, columnspan=2, pady=5, sticky="ew")
            
            # Message body
            self.body_label = ttk.Label(self.main_frame, text="Message:")
            self.body_label.grid(row=6, column=0, sticky="nw", pady=5)
            self.body_text = tk.Text(self.main_frame, width=35, height=7, font=("Segoe UI", 10), bd=2, relief="flat")
            self.body_text.grid(row=6, column=1, columnspan=2, pady=5, sticky="ew")
            
            # Attachment
            self.attachment_label = ttk.Label(self.main_frame, text="Attachment: None")
            self.attachment_label.grid(row=7, column=0, sticky="w", pady=5)
            self.attach_button = ttk.Button(self.main_frame, text="Add Attachment", command=self.add_attachment, style="Accent.TButton")
            self.attach_button.grid(row=7, column=1, pady=5, sticky="ew")
            self.clear_attach_button = ttk.Button(self.main_frame, text="Clear", command=self.clear_attachment)
            self.clear_attach_button.grid(row=7, column=2, pady=5, sticky="e")
            
            # Template and history
            self.template_label = ttk.Label(self.main_frame, text="Select Template/History:")
            self.template_label.grid(row=8, column=0, sticky="w", pady=5)
            self.template_combobox = ttk.Combobox(self.main_frame, values=["None"] + list(self.templates.keys()), width=20)
            self.template_combobox.grid(row=8, column=1, pady=5, sticky="ew")
            self.template_combobox.bind("<<ComboboxSelected>>", self.load_template)
            self.save_template_button = ttk.Button(self.main_frame, text="Save as Template", command=self.save_template)
            self.save_template_button.grid(row=8, column=2, pady=5, sticky="e")
            
            # Send and theme buttons
            self.send_button = ttk.Button(self.main_frame, text="Send Email", command=self.send_email, style="Accent.TButton")
            self.send_button.grid(row=9, column=0, columnspan=2, pady=10, sticky="ew")
            self.toggle_theme_button = ttk.Button(self.main_frame, text="Toggle Dark Mode", command=self.toggle_theme)
            self.toggle_theme_button.grid(row=9, column=2, pady=10, sticky="e")
            
            # Apply styles after all widgets are created
            self.configure_styles()
            
            # Center the window
            self.root.eval('tk::PlaceWindow . center')
        
        except Exception as e:
            with open("error.log", "w") as f:
                f.write(f"Error during initialization: {str(e)}\n")
                f.write(traceback.format_exc())
            messagebox.showerror("Error", "Failed to launch application. Check error.log for details.")
            self.root.destroy()
    
    def configure_styles(self):
        try:
            if self.dark_mode:
                bg_color = "#1e1e1e"
                fg_color = "#ffffff"
                entry_bg = "#2d2d2d"
                accent_color = "#6200ea"
            else:
                bg_color = "#e8ecef"
                fg_color = "#212121"
                entry_bg = "#ffffff"
                accent_color = "#3f51b5"
            
            self.style.configure("TLabel", font=("Segoe UI", 10), background=bg_color, foreground=fg_color)
            self.style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6, background=bg_color, foreground=fg_color)
            self.style.configure("Accent.TButton", background=accent_color, foreground="#ffffff")
            self.style.map("Accent.TButton",
                        background=[('active', '#536dfe'), ('pressed', '#304ffe')],
                        foreground=[('active', '#ffffff'), ('pressed', '#ffffff')])
            self.style.configure("TEntry", font=("Segoe UI", 10), padding=5, fieldbackground=entry_bg)
            self.style.configure("TCombobox", font=("Segoe UI", 10), padding=5, fieldbackground=entry_bg)
            self.style.configure("Main.TFrame", background=bg_color)
            self.root.configure(bg=bg_color)
            if hasattr(self, 'body_text'):
                self.body_text.configure(bg=entry_bg, fg=fg_color, insertbackground=fg_color)
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in configure_styles: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def toggle_theme(self):
        try:
            self.dark_mode = not self.dark_mode
            self.configure_styles()
            self.toggle_theme_button.configure(text="Toggle Light Mode" if self.dark_mode else "Toggle Dark Mode")
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in toggle_theme: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def load_credentials(self):
        try:
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, 'r') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in load_credentials: {str(e)}\n")
                f.write(traceback.format_exc())
            return {}
    
    def save_credentials(self):
        try:
            if self.save_credentials_var.get():
                credentials = {
                    "email": self.sender_entry.get(),
                    "password": self.password_entry.get()
                }
                with open(self.credentials_file, 'w') as f:
                    json.dump(credentials, f, indent=4)
            else:
                if os.path.exists(self.credentials_file):
                    os.remove(self.credentials_file)
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in save_credentials: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def add_attachment(self):
        try:
            file_path = filedialog.askopenfilename()
            if file_path:
                self.attachment_path = file_path
                self.attachment_label.configure(text=f"Attachment: {os.path.basename(file_path)}")
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in add_attachment: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def clear_attachment(self):
        try:
            self.attachment_path = None
            self.attachment_label.configure(text="Attachment: None")
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in clear_attachment: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def load_templates(self):
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in load_templates: {str(e)}\n")
                f.write(traceback.format_exc())
            return {}
    
    def save_template(self):
        try:
            subject = self.subject_entry.get()
            body = self.body_text.get("1.0", tk.END).strip()
            recipients = self.recipient_entry.get()
            if not subject or not body:
                messagebox.showerror("Error", "Subject and message are required to save a template.")
                return
            template_name = f"Template {len(self.templates) + 1}"
            self.templates[template_name] = {
                "subject": subject,
                "body": body,
                "recipients": recipients,
                "attachment": os.path.basename(self.attachment_path) if self.attachment_path else None
            }
            with open(self.history_file, 'w') as f:
                json.dump(self.templates, f, indent=4)
            self.template_combobox['values'] = ["None"] + list(self.templates.keys())
            messagebox.showinfo("Success", f"Template '{template_name}' saved!")
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in save_template: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def load_template(self, event):
        try:
            selected = self.template_combobox.get()
            if selected != "None" and selected in self.templates:
                self.subject_entry.delete(0, tk.END)
                self.subject_entry.insert(0, self.templates[selected]["subject"])
                self.body_text.delete("1.0", tk.END)
                self.body_text.insert("1.0", self.templates[selected]["body"])
                self.recipient_entry.delete(0, tk.END)
                self.recipient_entry.insert(0, self.templates[selected]["recipients"])
                self.attachment_path = None
                self.attachment_label.configure(text=f"Attachment: {self.templates[selected]['attachment'] or 'None'}")
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in load_template: {str(e)}\n")
                f.write(traceback.format_exc())
    
    def send_email(self):
        try:
            sender_email = self.sender_entry.get()
            app_password = self.password_entry.get()
            recipient_emails = [email.strip() for email in self.recipient_entry.get().split(",") if email.strip()]
            subject = self.subject_entry.get()
            body = self.body_text.get("1.0", tk.END).strip()
            
            # Save credentials if checked
            self.save_credentials()
            
            # Input validation
            if not sender_email or not app_password or not recipient_emails or not subject or not body:
                messagebox.showerror("Error", "Please fill in all required fields.")
                return
            
            # Email setup
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipient_emails)
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            
            # Add attachment if present
            if self.attachment_path:
                try:
                    with open(self.attachment_path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={os.path.basename(self.attachment_path)}'
                    )
                    msg.attach(part)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to attach file: {str(e)}")
                    return
            
            # Connect to Gmail's SMTP server
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, app_password)
            server.sendmail(sender_email, recipient_emails, msg.as_string())
            server.quit()
            
            # Save to history
            self.templates[f"History {len(self.templates) + 1}"] = {
                "subject": subject,
                "body": body,
                "recipients": ", ".join(recipient_emails),
                "attachment": os.path.basename(self.attachment_path) if self.attachment_path else None
            }
            with open(self.history_file, 'w') as f:
                json.dump(self.templates, f, indent=4)
            self.template_combobox['values'] = ["None"] + list(self.templates.keys())
            
            messagebox.showinfo("Success", "Email sent successfully!")
            
            # Clear fields
            self.recipient_entry.delete(0, tk.END)
            self.subject_entry.delete(0, tk.END)
            self.body_text.delete("1.0", tk.END)
            self.clear_attachment()
            
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(f"Error in send_email: {str(e)}\n")
                f.write(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")

def main():
    try:
        root = tk.Tk()
        app = EmailSenderApp(root)
        root.mainloop()
    except Exception as e:
        with open("error.log", "w") as f:
            f.write(f"Error in main: {str(e)}\n")
            f.write(traceback.format_exc())

if __name__ == "__main__":
    main()