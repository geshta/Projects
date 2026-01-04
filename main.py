import sys
import customtkinter as ctk
from tkinter import messagebox
from tabs.customer_tab import CustomerTab
from tabs.entry_tab import EntryTab
from tabs.message_tab import MessageTab
from tabs.reports_tab import ReportTab
from app_config import app_config
from setup_dialog import SetupDialog


class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Milk Delivery App")
        self.state("zoomed")
        self.configure(fg_color="#ffffff")

        # Set icon if available
        try:
            icon_path = app_config.app_dir / "MDA_LOGO.ico"
            if icon_path.exists():
                self.iconbitmap(str(icon_path))
        except:
            pass

        self.colors = {
            "primary": "#143583",
            "secondary": "#2563eb",
            "success": "#059669",
            "danger": "#e11d48",
            "warning": "#f59e0b",
            "info": "#0ea5e9",
            "bg_dark": "#1e1e22",
            "bg_light": "#f9fafb",
            "text_dark": "#111827",
            "text_light": "#f1f5f9",
            "accent": "#fbbf24",
            "milk": "#f0f9ff",
        }

        self.container = ctk.CTkFrame(self, fg_color=self.colors["bg_light"])
        self.container.pack(fill="both", expand=True)

        # Check if first run and show setup dialog
        if app_config.is_first_run():
            self.after(500, self.show_setup_dialog)

        # Initialize tabs
        try:
            self.customer_tab = CustomerTab(
                self.container, colors=self.colors, back_callback=self.show_home)
            self.entry_tab = EntryTab(
                self.container, colors=self.colors, back_callback=self.show_home, 
                customer_tab=self.customer_tab)
            self.message_tab = MessageTab(
                self.container, colors=self.colors, back_callback=self.show_home, 
                customer_tab=self.customer_tab)
            self.report_tab = ReportTab(
                self.container, colors=self.colors, back_callback=self.show_home)
            self.customer_tab.on_customer_change = self.refresh_all_customer_data

        except Exception as e:
            print(f"Error initializing tabs: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "Initialization Error",
                f"Failed to initialize application:\n{e}\n\nPlease contact support."
            )
            sys.exit(1)

        # Initially hide all tabs
        self.customer_tab.pack_forget()
        self.entry_tab.pack_forget()
        self.message_tab.pack_forget()
        self.report_tab.pack_forget()

        # Show splash screen
        # Show splash screen
        self.show_splash_screen()
        self.tree = None
        self.after(0, lambda: self.state("zoomed"))
    
    
        
    
    def _ui_call(self, func):
        """Run UI code safely from any thread"""
        try:
            self.after(0, func)
        except RuntimeError:
            pass


    
    def _begin_container_build(self):
        """Hide main container while building a page"""
        self.container.pack_forget()
        self.update_idletasks()


    def _end_container_build(self):
        """Show container after page is fully built"""
        self.container.pack(fill="both", expand=True)
        self.update_idletasks()




    
    def refresh_all_customer_data(self):
        """Refresh customer data across all tabs when customer info changes"""
        try:
            # 1. Reload Entry Tab
            if hasattr(self, 'entry_tab') and self.entry_tab:
                try:
                    # Entry tab doesn't have customer list, but force any cached data to refresh
                    pass
                except Exception as e:
                    print(f"Entry tab refresh error: {e}")
            
            # 2. Reload Message Tab
            if hasattr(self, 'message_tab') and self.message_tab:
                try:
                    # If message tab has customer data loaded, reload it
                    if hasattr(self.message_tab, 'customer_data') and self.message_tab.customer_data:
                        if hasattr(self.message_tab, 'selected_month') and self.message_tab.selected_month:
                            # Reload the current month's data
                            self.message_tab.load_customer_list(self.message_tab.selected_month)
                except Exception as e:
                    print(f"Message tab refresh error: {e}")
            
            # 3. Reload Report Tab
            if hasattr(self, 'report_tab') and self.report_tab:
                try:
                    # Reload report data based on current view mode
                    if hasattr(self.report_tab, 'view_mode'):
                        if self.report_tab.view_mode == "all_records":
                            # Refresh all records view
                            if hasattr(self.report_tab, 'refresh_all_records'):
                                self.report_tab.refresh_all_records()
                        elif self.report_tab.view_mode == "monthly" and hasattr(self.report_tab, 'selected_month'):
                            # Refresh monthly view
                            self.report_tab.show_monthly_data(self.report_tab.selected_month)
                        elif self.report_tab.view_mode == "yearly" and hasattr(self.report_tab, 'selected_year'):
                            # Refresh yearly view
                            self.report_tab.show_yearly_data(self.report_tab.selected_year)
                except Exception as e:
                    print(f"Report tab refresh error: {e}")
                    
        except Exception as e:
            print(f"Global refresh error: {e}")



    def show_setup_dialog(self):
        """Show setup dialog on first run"""
        setup = SetupDialog(self)
        self.wait_window(setup)

    def show_splash_screen(self):
        """Show welcome splash screen"""
        splash_frame = ctk.CTkFrame(self.container, fg_color="#ffffff")
        splash_frame.pack(fill="both", expand=True)

        # Create gradient background
        top_gradient = ctk.CTkFrame(splash_frame, fg_color="#0ea5e9", height=200)
        top_gradient.pack(fill="x")
        top_gradient.pack_propagate(False)

        # Milk icon area
        milk_icon_frame = ctk.CTkFrame(splash_frame, fg_color="#ffffff")
        milk_icon_frame.pack(fill="both", expand=True)

        # Large welcome emoji
        ctk.CTkLabel(
            milk_icon_frame,
            text="🥛",
            font=ctk.CTkFont(size=120),
            text_color="#0ea5e9"
        ).pack(pady=(40, 20))

        # Welcome text
        user_name = app_config.get_user_name()
        welcome_label = ctk.CTkLabel(
            milk_icon_frame,
            text=f"Welcome, {user_name}! 👋",
            font=ctk.CTkFont(size=48, weight="bold"),
            text_color=self.colors["primary"]
        )
        welcome_label.pack(pady=10)

        # Subtitle
        ctk.CTkLabel(
            milk_icon_frame,
            text="Your Milk Delivery Management System",
            font=ctk.CTkFont(size=18),
            text_color="#6b7280"
        ).pack(pady=(5, 50))

        # Loading indicator
        loading_frame = ctk.CTkFrame(milk_icon_frame, fg_color="transparent")
        loading_frame.pack(pady=30)

        dots = ["●  ○  ○", "●  ●  ○", "●  ●  ●"]
        dot_label = ctk.CTkLabel(
            loading_frame,
            text=dots[0],
            font=ctk.CTkFont(size=20),
            text_color=self.colors["secondary"]
        )
        dot_label.pack()

        # Animate dots
        def animate_dots(index=0):
            if index < 3:
                dot_label.configure(text=dots[index])
                milk_icon_frame.after(300, lambda: animate_dots(index + 1))
            else:
                milk_icon_frame.after(500, lambda: self.transition_to_home(splash_frame))

        milk_icon_frame.after(500, animate_dots)

    def transition_to_home(self, splash_frame):
        splash_frame.pack_forget()
        self.show_home()


    
        
    

    
    def show_home(self):
        """Show home screen"""
        self._begin_container_build()   # ⬅ ADD
        # After header creation
        


        try:
            self.customer_tab.pack_forget()
        except:
            pass
        
        try:
            self.entry_tab.pack_forget()
        except:
            pass
        
        try:
            self.message_tab.pack_forget()
        except:
            pass
        
        try:
            self.report_tab.pack_forget()
        except:
            pass
        
        try:
            self.account_frame.pack_forget()
        except:
            pass
        
        

        # Create home frame if not exists
        if hasattr(self, 'home_frame') and self.home_frame.winfo_exists():
            self.home_frame.pack(fill="both", expand=True)
            self._end_container_build()
            return
        

            

        self.home_frame = ctk.CTkFrame(self.container, fg_color="#ffffff")
        self.home_frame.pack(fill="both", expand=True)

        


        # Header
        header = ctk.CTkFrame(self.home_frame, fg_color=self.colors["primary"], height=120)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Header content
        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=30, pady=20)

        # Welcome message
        user_name = app_config.get_user_name()
        welcome_frame = ctk.CTkFrame(header_inner, fg_color="transparent")
        welcome_frame.pack(side="left", fill="both", expand=True)
        
        ctk.CTkLabel(
            welcome_frame,
            text=f"Welcome back, {user_name}! 🥛",
            font=ctk.CTkFont(size=44, weight="bold"),
            text_color="#ffffff"
        ).pack(anchor="w")

        # Business name
        business_name = app_config.get_business_name()
        ctk.CTkLabel(
            welcome_frame,
            text=f"{business_name}",
            font=ctk.CTkFont(size=16),
            text_color="#e0f2fe"
        ).pack(anchor="w", pady=(5, 0))

        # Account button (top right)
        account_btn = ctk.CTkButton(
            header_inner,
            text="👤",
            font=ctk.CTkFont(size=32),
            width=70,
            height=70,
            corner_radius=35,
            fg_color="#ffffff",
            text_color=self.colors["primary"],
            hover_color="#e0f2fe",
            command=self.show_account_settings
        )
        account_btn.pack(side="right")

        # Main content area
        content_frame = ctk.CTkFrame(self.home_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=40, pady=40)

        # Grid layout for buttons
        button_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        button_frame.pack(fill="both", expand=True)

        buttons_info = [
            {
                "text": "👥 CUSTOMERS",
                "color": "#3b82f6",
                "hover": "#1d4ed8",
                "command": self.show_customer_tab,
                "description": "Manage customer database"
            },
            {
                "text": "📝 ENTRIES",
                "color": "#143583",
                "hover": "#0f2460",
                "command": self.show_entry_tab,
                "description": "Record daily deliveries"
            },
            {
                "text": "💬 MESSAGES",
                "color": "#059669",
                "hover": "#047857",
                "command": self.show_message_tab,
                "description": "Send WhatsApp messages"
            },
            {
                "text": "📊 REPORTS",
                "color": "#f59e0b",
                "hover": "#d97706",
                "command": self.show_report_tab,
                "description": "View customer reports"
            },
        ]

        

        for idx, btn_info in enumerate(buttons_info):
            row = idx // 2
            col = idx % 2

            # Button container
            btn_container = ctk.CTkFrame(button_frame, fg_color="transparent")
            btn_container.grid(row=row, column=col, padx=15, pady=15, sticky="nsew")

            # Button
            btn = ctk.CTkButton(
                btn_container,
                text=btn_info["text"],
                width=300,
                height=100,
                fg_color=btn_info["color"],
                text_color="white",
                font=ctk.CTkFont(size=28, weight="bold"),
                corner_radius=20,
                command=btn_info["command"],
                hover_color=btn_info["hover"]
            )
            btn.pack(fill="both", expand=True)

            # Description
            desc_label = ctk.CTkLabel(
                btn_container,
                text=btn_info["description"],
                font=ctk.CTkFont(size=11),
                text_color="#6b7280"
            )
            desc_label.pack(pady=(8, 0))


            # Right side buttons container
            right_buttons = ctk.CTkFrame(header_inner, fg_color="transparent")
            right_buttons.pack(side="right")

            
        account_btn = ctk.CTkButton(
            right_buttons,
            text="👤",
            font=ctk.CTkFont(size=32),
            width=70,
            height=70,
            corner_radius=35,
            fg_color="#ffffff",
            text_color=self.colors["primary"],
            hover_color="#e0f2fe",
            command=self.show_account_settings
            )
        account_btn.pack(side="left")


        # Configure grid weights
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_rowconfigure(0, weight=1)
        button_frame.grid_rowconfigure(1, weight=1)

        # Footer
        footer = ctk.CTkFrame(self.home_frame, fg_color="#f3f4f6", height=80)
        footer.pack(fill="x")
        footer.pack_propagate(False)

        footer_inner = ctk.CTkFrame(footer, fg_color="transparent")
        footer_inner.pack(fill="both", expand=True, padx=40, pady=15)

        ctk.CTkLabel(
            footer_inner,
            text="✨ Version 1.0 | Milk Delivery Management System",
            font=ctk.CTkFont(size=12),
            text_color="#6b7280"
        ).pack(side="left")

        ctk.CTkLabel(
            footer_inner,
            text="🔒 Your data is secure and encrypted",
            font=ctk.CTkFont(size=12),
            text_color="#059669"
        ).pack(side="right")
        
        self._end_container_build()     # ⬅ ADD
    

   




    def show_account_settings(self):
        """Show account settings page with live message preview"""
        self._begin_container_build() 
        try:
            self.home_frame.pack_forget()
        except:
            pass
        # Clear previous account UI to avoid duplicate buttons
        if hasattr(self, "account_frame") and self.account_frame.winfo_exists():
            for widget in self.account_frame.winfo_children():
                widget.destroy()


        self.account_frame = ctk.CTkFrame(self.container, fg_color="#ffffff")
        self.account_frame.pack(fill="both", expand=True)

        # ================= HEADER =================
        header = ctk.CTkFrame(self.account_frame, fg_color=self.colors["primary"], height=100)
        header.pack(fill="x")
        header.pack_propagate(False)

        
        
        


        ctk.CTkButton(
            header,
            text="← Back",
            font=ctk.CTkFont(size=20, weight="bold"),
            width=140,
            height=60,
            corner_radius=15,
            fg_color=self.colors["text_light"],
            text_color=self.colors["primary"],
            hover_color="#e2e8f0",
            command=self.show_home,
        ).pack(side="left", padx=30, pady=20)

        ctk.CTkLabel(
            header,
            text="👤 Account Settings",
            font=ctk.CTkFont(size=40, weight="bold"),
            text_color=self.colors["text_light"],
        ).pack(side="left", padx=20)

        # ================= BODY =================
        body = ctk.CTkFrame(self.account_frame, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=40, pady=30)

        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)

        # ================= LEFT: FORM =================
        content = ctk.CTkScrollableFrame(
            body,
            fg_color="transparent",
            scrollbar_button_color=self.colors["primary"]
        )
        content.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

        # 🔧 SCROLL SPEED FIX
        content.bind_all(
            "<MouseWheel>",
            lambda e: content._parent_canvas.yview_scroll(int(-1 * (e.delta / 10)), "units")
        )


        self.create_editable_field(content, "👤 Enter your Name", "user_name", app_config.get_user_name())
        self.create_editable_field(content, "🏢 Business/Dairy Name", "business_name", app_config.get_business_name())
        self.create_editable_field(content, "📞 Contact Phone Number", "contact_number", app_config.get_contact_number())
        self.create_editable_field(content, "💳 Payment Methods", "payment_info", app_config.get_payment_info())

        # ================= RIGHT: MESSAGE PREVIEW =================
        preview_frame = ctk.CTkFrame(
            body,
            fg_color="#f8fafc",
            corner_radius=18,
            border_width=2,
            border_color="#e5e7eb"
        )
        preview_frame.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(
            preview_frame,
            text="📱 WhatsApp Message Preview",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=self.colors["primary"]
        ).pack(pady=15)

        self.preview_label = ctk.CTkLabel(
            preview_frame,
            text="",
            font=ctk.CTkFont(size=14),
            text_color="#111827",
            justify="left",
            anchor="nw",
            wraplength=360
        )
        self.preview_label.pack(fill="both", expand=True, padx=20, pady=10)

        # Initial preview
        self.update_message_preview()

        # ================= ACTIONS =================
        action_frame = ctk.CTkFrame(self.account_frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=(10, 30))

        ctk.CTkButton(
            action_frame,
            text="💾 Save Changes",
            font=ctk.CTkFont(size=20, weight="bold"),
            width=250,
            height=70,
            corner_radius=15,
            fg_color="#059669",
            hover_color="#047857",
            command=self.save_account_settings
        ).pack(side="left", padx=15)

        ctk.CTkButton(
            action_frame,
            text="🔄 Reset",
            font=ctk.CTkFont(size=20, weight="bold"),
            width=200,
            height=70,
            corner_radius=15,
            fg_color="#6b7280",
            hover_color="#4b5563",
            command=self.reset_account_fields
        ).pack(side="left")

        # Bind live preview updates
        for field in ["user_name", "business_name", "contact_number", "payment_info"]:
            entry = getattr(self, f"account_entry_{field}")
            entry.bind("<KeyRelease>", lambda e: self.update_message_preview())
            
        self._end_container_build() 


    def update_message_preview(self):
        """Live WhatsApp message preview"""
        name = self.account_entry_user_name.get()
        business = self.account_entry_business_name.get()
        phone = self.account_entry_contact_number.get()
        payment = self.account_entry_payment_info.get()

        preview = (
            f"🥛 {business or 'Your Dairy'} - Monthly Bill\n\n"
            f"Dear Customer,\n\n"
            f"📅 Period: 01 JAN 2026 - 31 JAN 2026\n"
            f"📊 Delivery Summary:\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"🥛 Total Milk: 45.00 Liters\n"
            f"💰 Amount Due: ₹2250.00\n"
            f"━━━━━━━━━━━━━━━━━━\n\n"
            f"📞 For queries: {phone or 'XXXXXXXXXX'}\n"
            f"💳 Pay via: {payment or 'UPI / Cash'}\n\n"
            f"Thank you!\n"
            f"— {name or 'Owner'}"
        )

        self.preview_label.configure(text=preview)


    
    def create_editable_field(self, parent, label_text, field_name, current_value):
        """Create an editable field in account settings"""
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=(0, 25))
        
        # Label
        ctk.CTkLabel(
            field_frame,
            text=label_text,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#1f2937",
            anchor="w"
        ).pack(fill="x", pady=(0, 10))
        
        # Entry field
        entry = ctk.CTkEntry(
            field_frame,
            font=ctk.CTkFont(size=16),
            height=55,
            corner_radius=12,
            border_width=2,
            border_color="#cbd5e1",
            fg_color="#ffffff"
        )
        entry.insert(0, current_value)
        entry.pack(fill="x")
        
        # Store reference
        setattr(self, f"account_entry_{field_name}", entry)
    
    def save_account_settings(self):
        """Save account settings"""
        user_name = self.account_entry_user_name.get().strip()
        business_name = self.account_entry_business_name.get().strip()
        contact_number = self.account_entry_contact_number.get().strip()
        payment_info = self.account_entry_payment_info.get().strip()
        
        # Validate
        if not user_name:
            messagebox.showerror("Validation Error", "Please enter your name")
            return
        
        if len(user_name) < 2:
            messagebox.showerror("Validation Error", "Name must be at least 2 characters long")
            return
        
        if not business_name:
            messagebox.showerror("Validation Error", "Please enter your business name")
            return
        
        if not contact_number:
            messagebox.showerror("Validation Error", "Please enter contact number")
            return
        
        clean_number = ''.join(filter(str.isdigit, contact_number))
        if len(clean_number) != 10:
            messagebox.showerror(
                "Validation Error",
                f"Please enter a valid 10-digit phone number\n\nYou entered: {contact_number}\nDigits found: {len(clean_number)}"
            )
            return
        
        if not payment_info:
            messagebox.showerror("Validation Error", "Please enter payment information")
            return
        
        # Save
        app_config.update_settings(
            user_name=user_name,
            business_name=business_name,
            contact_number=contact_number,
            payment_info=payment_info
        )
        
        messagebox.showinfo(
            "Settings Saved",
            "✅ Your account settings have been saved successfully!"
        )
        
        # Refresh home screen if it exists
        if hasattr(self, 'home_frame'):
            self.home_frame.destroy()
            delattr(self, 'home_frame')
    
    def reset_account_fields(self):
        """Reset fields to current saved values"""
        self.account_entry_user_name.delete(0, "end")
        self.account_entry_user_name.insert(0, app_config.get_user_name())
        
        self.account_entry_business_name.delete(0, "end")
        self.account_entry_business_name.insert(0, app_config.get_business_name())
        
        self.account_entry_contact_number.delete(0, "end")
        self.account_entry_contact_number.insert(0, app_config.get_contact_number())
        
        self.account_entry_payment_info.delete(0, "end")
        self.account_entry_payment_info.insert(0, app_config.get_payment_info())
        
        messagebox.showinfo("Reset", "Fields have been reset to saved values")

    def show_customer_tab(self):
        """Show customer tab"""
        try:
            self.home_frame.pack_forget()
        except:
            pass
        try:
            self.entry_tab.pack_forget()
        except:
            pass
        try:
            self.message_tab.pack_forget()
        except:
            pass
        try:
            self.report_tab.pack_forget()
        except:
            pass
        
        self.customer_tab.pack(fill="both", expand=True)

    def show_entry_tab(self):
        """Show entry tab"""
        try:
            self.home_frame.pack_forget()
        except:
            pass
        try:
            self.customer_tab.pack_forget()
        except:
            pass
        try:
            self.message_tab.pack_forget()
        except:
            pass
        try:
            self.report_tab.pack_forget()
        except:
            pass
        
        self.entry_tab.pack(fill="both", expand=True)

    def show_message_tab(self):
        """Show message tab"""
        try:
            self.home_frame.pack_forget()
        except:
            pass
        try:
            self.customer_tab.pack_forget()
        except:
            pass
        try:
            self.entry_tab.pack_forget()
        except:
            pass
        try:
            self.report_tab.pack_forget()
        except:
            pass
        
        self.message_tab.pack(fill="both", expand=True)

    def show_report_tab(self):
        """Show report tab"""
        try:
            self.home_frame.pack_forget()
        except:
            pass
        try:
            self.customer_tab.pack_forget()
        except:
            pass
        try:
            self.entry_tab.pack_forget()
        except:
            pass
        try:
            self.message_tab.pack_forget()
        except:
            pass
        
        try:
            self.report_tab.pack(fill="both", expand=True)
        except Exception as e:
            print(f"Error showing report tab: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")
    
    try:
        app = MainApp()
        app.mainloop()
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Fatal Error", f"Application crashed:\n{e}")