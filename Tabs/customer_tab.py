import customtkinter as ctk
from tkinter import ttk, messagebox
import pandas as pd
import os
import json
import datetime  


class CustomerTab(ctk.CTkFrame):
    # Replace the __init__ method in CustomerTab class with this:

    def __init__(self, parent, colors, back_callback):
        
        super().__init__(parent)
        self.parent = parent
        self.colors = colors
        self.back_callback = back_callback

        # UPDATED: Use dynamic paths from app_config
        from app_config import app_config
        self.app_config = app_config
        
        self.excel_file = str(app_config.customers_file)
        self.deleted_file = str(app_config.deleted_file)
        self.cluster_file = str(app_config.cluster_file)

        self.cluster_bg_colors = [
            "#6ee7b7",
            "#facc15",
            "#60a5fa",
            "#f87171",
            "#93c5fd",
            "#fdba74",
            "#34d399",
        ]

        self.on_customer_change = None
        
        # Undo functionality
        self.undo_stack = []
        self.max_undo = 10

        self.load_clusters()
        self.initialize_excel()
        self.initialize_deleted_excel()
        self.pack_forget()
        self.create_ui()
        self.load_data()
    
    def _begin_page_build(self):
        """Hide frame while building page"""
        self.pack_forget()
        self.update_idletasks()

    def _end_page_build(self):
        """Show frame after page is built"""
        self.pack(fill="both", expand=True)
        self.update_idletasks()

    def load_clusters(self):
        if not os.path.exists(self.cluster_file):
            self.clusters = [
                {"id": 1, "name": "Cluster 1", "morn": 1.0, "even": 1.0},
                {"id": 2, "name": "Cluster 2", "morn": 2.5, "even": 2.0},
            ]
            with open(self.cluster_file, "w") as f:
                json.dump(self.clusters, f)
        else:
            with open(self.cluster_file, "r") as f:
                self.clusters = json.load(f)
            changed = False
            for idx, c in enumerate(self.clusters, start=1):
                if "name" not in c:
                    c["name"] = f"Cluster {c.get('id', idx)}"
                    changed = True
            if changed:
                with open(self.cluster_file, "w") as f:
                    json.dump(self.clusters, f)

    def save_clusters(self):
        with open(self.cluster_file, "w") as f:
            json.dump(self.clusters, f)

    def initialize_excel(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
            df.to_excel(self.excel_file, index=False, engine="openpyxl")
        else:
            try:
                df = pd.read_excel(self.excel_file, engine="openpyxl")
                columns = df.columns.tolist()
                if df.empty or columns != ["S.No", "CID", "Name", "Phone", "Address", "Cluster"]:
                    df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
                    df.to_excel(self.excel_file, index=False, engine="openpyxl")
            except:
                df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
                df.to_excel(self.excel_file, index=False, engine="openpyxl")

    def initialize_deleted_excel(self):
        if not os.path.exists(self.deleted_file):
            df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
            df.to_excel(self.deleted_file, index=False, engine="openpyxl")
        else:
            try:
                df = pd.read_excel(self.deleted_file, engine="openpyxl")
                columns = df.columns.tolist()
                if df.empty or columns != ["S.No", "CID", "Name", "Phone", "Address", "Cluster"]:
                    df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
                    df.to_excel(self.deleted_file, index=False, engine="openpyxl")
            except:
                df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])
                df.to_excel(self.deleted_file, index=False, engine="openpyxl")

    def create_ui(self):
        header = ctk.CTkFrame(self, fg_color=self.colors["primary"], height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        btn_back = ctk.CTkButton(
            header,
            text="Back",
            font=ctk.CTkFont(size=24, weight="bold"),
            width=150,
            height=55,
            corner_radius=15,
            fg_color=self.colors["text_light"],
            text_color=self.colors["primary"],
            hover_color="#e2e8f0",
            command=self.back_callback,
        )
        btn_back.pack(side="left", padx=30, pady=12)

        lbl_title = ctk.CTkLabel(
            header,
            text="👥 Customer Details Management",
            font=ctk.CTkFont(size=40, weight="bold"),
            text_color=self.colors["text_light"],
        )
        lbl_title.pack(side="left", padx=20)

        form_row = ctk.CTkFrame(self, fg_color="transparent")
        form_row.pack(fill="x", padx=40, pady=30)

        frame_add = ctk.CTkFrame(
            form_row,
            fg_color=self.colors["text_light"],
            corner_radius=15,
            border_width=2,
            border_color=self.colors["success"],
        )
        frame_add.pack(fill="x", pady=(0, 15))

        lbl_add_title = ctk.CTkLabel(
            frame_add,
            text="➕ Add Customer:",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=self.colors["success"],
        )
        lbl_add_title.pack(side="left", padx=20, pady=15)

        self.lbl_cid = ctk.CTkLabel(
            frame_add,
            text=self.get_next_cid(),
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#f0fdf4",
            corner_radius=8,
            width=70,
            height=40,
        )
        self.lbl_cid.pack(side="left", padx=10)

        self.ent_name = ctk.CTkEntry(
            frame_add,
            placeholder_text="Name",
            font=ctk.CTkFont(size=16),
            height=40,
            width=220,
            corner_radius=8,
            border_width=2,
            border_color=self.colors["primary"],
            fg_color="#ffffff",
            text_color="#000000",
        )
        self.ent_name.pack(side="left", padx=(0, 10))

        self.ent_phone = ctk.CTkEntry(
            frame_add,
            placeholder_text="Phone (10 digits)",
            font=ctk.CTkFont(size=16),
            width=160,
            height=40,
            corner_radius=8,
            border_width=2,
            border_color=self.colors["primary"],
            fg_color="#ffffff",
            text_color="#000000",
        )
        self.ent_phone.pack(side="left", padx=(0, 10))

        self.ent_address = ctk.CTkEntry(
            frame_add,
            placeholder_text="Address",
            font=ctk.CTkFont(size=16),
            height=40,
            width=400,
            corner_radius=8,
            border_width=2,
            border_color=self.colors["primary"],
            fg_color="#ffffff",
            text_color="#000000",
        )
        self.ent_address.pack(side="left", padx=(0, 10))

        btn_add = ctk.CTkButton(
            frame_add,
            text="Add",
            font=ctk.CTkFont(size=16, weight="bold"),
            width=100,
            height=40,
            corner_radius=8,
            fg_color=self.colors["success"],
            hover_color="#059669",
            command=self.add_customer,
        )
        btn_add.pack(side="left")

        frame_list = ctk.CTkFrame(self, fg_color=self.colors["text_light"])
        frame_list.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        list_header = ctk.CTkFrame(frame_list, fg_color="transparent", height=70)
        list_header.pack(fill="x")

        lbl_list_title = ctk.CTkLabel(
            list_header,
            text="📋 Customer List",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=self.colors["primary"],
        )
        lbl_list_title.pack(side="left")

        frm_actions = ctk.CTkFrame(list_header, fg_color="transparent")
        frm_actions.pack(side="right")

        self.ent_search = ctk.CTkEntry(
            frm_actions,
            placeholder_text="Search CID, Name, Phone, Address",
            font=ctk.CTkFont(size=16),
            width=300,
            height=45,
            corner_radius=15,
            border_width=2,
            border_color=self.colors["info"],
            fg_color="#ffffff",
            text_color="#000000",
        )
        self.ent_search.pack(side="left", padx=(0, 10))

        btn_search = ctk.CTkButton(
            frm_actions,
            text="Search",
            font=ctk.CTkFont(size=18, weight="bold"),
            width=80,
            height=45,
            corner_radius=15,
            fg_color=self.colors["info"],
            hover_color="#0891b2",
            command=self.search_customers,
        )
        btn_search.pack(side="left", padx=(0, 10))

        btn_refresh = ctk.CTkButton(
            frm_actions,
            text="Refresh",
            font=ctk.CTkFont(size=18, weight="bold"),
            width=80,
            height=45,
            corner_radius=15,
            fg_color=self.colors["warning"],
            hover_color="#d97706",
            command=self.load_data,
        )
        btn_refresh.pack(side="left", padx=(0, 10))

        btn_edit = ctk.CTkButton(
            frm_actions,
            text="Edit",
            font=ctk.CTkFont(size=18, weight="bold"),
            width=80,
            height=45,
            corner_radius=15,
            fg_color="#8b5cf6",
            hover_color="#7c3aed",
            command=self.edit_customer,
        )
        btn_edit.pack(side="left", padx=(0, 10))

        btn_delete = ctk.CTkButton(
            frm_actions,
            text="Delete",
            font=ctk.CTkFont(size=18, weight="bold"),
            width=80,
            height=45,
            corner_radius=15,
            fg_color=self.colors["danger"],
            hover_color="#dc2626",
            command=self.delete_customers,
        )
        btn_delete.pack(side="left")

        table_container = ctk.CTkFrame(frame_list, fg_color="transparent")
        table_container.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background="#ffffff",
            foreground=self.colors["text_dark"],
            fieldbackground="#ffffff",
            borderwidth=0,
            rowheight=50,
            font=("Arial", 18),
        )
        style.configure(
            "Custom.Treeview.Heading",
            background=self.colors["primary"],
            foreground=self.colors["text_light"],
            borderwidth=0,
            relief="flat",
            font=("Arial", 20, "bold"),
        )
        style.map(
            "Custom.Treeview",
            background=[("selected", "#14b8a6")],
            foreground=[("selected", self.colors["text_light"])],
        )
        style.map(
            "Custom.Treeview.Heading", background=[("active", self.colors["secondary"])]
        )

        self.columns = ("S.No", "CID", "Name", "Phone", "Address")
        self.tree = ttk.Treeview(
        table_container,
        columns=self.columns,
        show="headings",
        selectmode="browse",   # IMPORTANT for keyboard navigation
        style="Custom.Treeview",
        )
        self.tree.bind("<Up>", self._on_arrow_up)
        self.tree.bind("<Down>", self._on_arrow_down)


        self.tree.pack(fill="both", expand=True)

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.tree.tag_configure("oddrow", background="#e0f2fe")
        self.tree.tag_configure("evenrow", background="#ffffff")

        self.tree.bind("<Button-1>", self.on_tree_click)

        btn_undo = ctk.CTkButton(
            frm_actions,
            text="↶ Undo",
            font=ctk.CTkFont(size=18, weight="bold"),
            width=80,
            height=45,
            corner_radius=15,
            fg_color="#f59e0b",
            hover_color="#d97706",
            command=self.undo_last_action,
        )
        btn_undo.pack(side="left", padx=(10, 0))
    
    def _on_arrow_up(self, event):
        self.tree.focus_set()
        items = self.tree.get_children()
        if not items:
            return "break"

        current = self.tree.focus() or items[0]
        index = items.index(current)

        if index > 0:
            item = items[index - 1]
            self.tree.focus(item)
            self.tree.selection_set(item)
            self.tree.see(item)

        return "break"


    def _on_arrow_down(self, event):
        self.tree.focus_set()
        items = self.tree.get_children()
        if not items:
            return "break"

        current = self.tree.focus() or items[0]
        index = items.index(current)

        if index < len(items) - 1:
            item = items[index + 1]
            self.tree.focus(item)
            self.tree.selection_set(item)
            self.tree.see(item)

        return "break"


    
    def _focus_first_row(self):
        try:
            items = self.tree.get_children()
            if items:
                self.tree.focus_set()
                self.tree.focus(items[0])
                self.tree.selection_set(items[0])
                self.tree.see(items[0])
        except:
            pass



    def get_next_cid(self):
        """Get next CID - NEVER reuse deleted CIDs"""
        try:
            all_cids = []
            
            # Get CIDs from active customers
            df_active = pd.read_excel(self.excel_file, engine="openpyxl")
            if not df_active.empty and "CID" in df_active.columns:
                for cid in df_active["CID"]:
                    if str(cid).startswith("C_"):
                        try:
                            all_cids.append(int(str(cid).split("_")[1]))
                        except:
                            continue
            
            # IMPORTANT: Also check deleted customers to avoid reusing CIDs
            try:
                df_deleted = pd.read_excel(self.deleted_file, engine="openpyxl")
                if not df_deleted.empty and "CID" in df_deleted.columns:
                    for cid in df_deleted["CID"]:
                        if str(cid).startswith("C_"):
                            try:
                                all_cids.append(int(str(cid).split("_")[1]))
                            except:
                                continue
            except:
                pass
            
            # Get maximum CID number ever used
            max_num = max(all_cids) if all_cids else 0
            return f"C_{max_num + 1}"
            
        except Exception as e:
            print(f"Error getting next CID: {e}")
            return "C_1"
        
    
    def save_undo_state(self, action_type, data):
        """Save current state for undo functionality"""
        try:
            state = {
                'action': action_type,
                'data': data.copy() if isinstance(data, pd.DataFrame) else data,
                'timestamp': datetime.datetime.now()
            }
            self.undo_stack.append(state)
            
            # Keep only last max_undo actions
            if len(self.undo_stack) > self.max_undo:
                self.undo_stack.pop(0)
        except Exception as e:
            print(f"Error saving undo state: {e}")
            
    
    def undo_last_action(self):
        """Undo the last delete/edit action"""
        if not self.undo_stack:
            messagebox.showinfo("Undo", "No actions to undo", parent=self)
            return
        
        try:
            last_state = self.undo_stack.pop()
            action = last_state['action']
            data = last_state['data']
            
            if action == 'delete':
                # Restore deleted customers
                df_active = pd.read_excel(self.excel_file, engine="openpyxl")
                df_deleted = pd.read_excel(self.deleted_file, engine="openpyxl")
                
                # Remove from deleted
                cids_to_restore = data['CID'].tolist()
                df_deleted = df_deleted[~df_deleted['CID'].isin(cids_to_restore)]
                
                # Add back to active
                df_active = pd.concat([df_active, data], ignore_index=True)
                df_active = df_active.sort_values('CID')
                df_active['S.No'] = range(1, len(df_active) + 1)
                
                # Save
                df_active.to_excel(self.excel_file, index=False, engine="openpyxl")
                df_deleted['S.No'] = range(1, len(df_deleted) + 1)
                df_deleted.to_excel(self.deleted_file, index=False, engine="openpyxl")
                
                self.load_data()
                messagebox.showinfo("Undo", f"Restored {len(cids_to_restore)} customer(s)", parent=self)
                
                if self.on_customer_change:
                    self.on_customer_change()
                    
            elif action == 'edit':
                # Restore previous values
                df = pd.read_excel(self.excel_file, engine="openpyxl")
                cid = data['CID']
                df.loc[df['CID'] == cid, 'Name'] = data['old_name']
                df.loc[df['CID'] == cid, 'Phone'] = data['old_phone']
                df.loc[df['CID'] == cid, 'Address'] = data['old_address']
                df.to_excel(self.excel_file, index=False, engine="openpyxl")
                
                self.load_data()
                messagebox.showinfo("Undo", f"Restored previous values for {cid}", parent=self)
                
                if self.on_customer_change:
                    self.on_customer_change()
                    
        except Exception as e:
            messagebox.showerror("Undo Error", f"Failed to undo: {e}", parent=self)

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.focus(item)
            self.tree.selection_set(item)




    def load_data(self):
        self.tree.delete(*self.tree.get_children())
        try:
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            if df.empty:
                return

            for idx, row in enumerate(df.itertuples(), start=0):
                tag = "oddrow" if idx % 2 else "evenrow"
                self.tree.insert(
                    "", "end",
                    values=(row._1, row.CID, row.Name, row.Phone, row.Address),
                    tags=(tag,)
                )

            self._focus_first_row()   # ← REQUIRED

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load customer data: {e}", parent=self)


    def add_customer(self):
        name = self.ent_name.get().strip()
        phone = self.ent_phone.get().strip()
        address = self.ent_address.get().strip()
        cluster_name = "Default"
        if not name:
            messagebox.showerror("Validation Error", "Please enter a name", parent=self)
            return
        if not phone or not phone.isdigit() or len(phone) != 10:
            messagebox.showerror("Validation Error", "Enter valid 10-digit phone number", parent=self)
            return
        if not address:
            messagebox.showerror("Validation Error", "Please enter address", parent=self)
            return
        try:
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            if phone in df["Phone"].astype(str).values:
                messagebox.showerror("Duplicate Entry", "Phone number already exists", parent=self)
                return
        except:
            df = pd.DataFrame(columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"])

        cid = self.get_next_cid()
        new_row = pd.DataFrame(
            [[None, cid, name, phone, address, cluster_name]],
            columns=["S.No", "CID", "Name", "Phone", "Address", "Cluster"],
        )
        try:
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            df = pd.concat([df, new_row], ignore_index=True)

            df["S.No"] = range(1, len(df) + 1)

            df.to_excel(self.excel_file, index=False, engine="openpyxl")
            self.ent_name.delete(0, "end")
            self.ent_phone.delete(0, "end")
            self.ent_address.delete(0, "end")
            self.lbl_cid.configure(text=self.get_next_cid())
            self.load_data()
            messagebox.showinfo("Success", f"Customer {cid} added successfully", parent=self)
            if self.on_customer_change:
                self.on_customer_change()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add customer: {e}", parent=self)

    def search_customers(self):
        term = self.ent_search.get().strip().lower()
        if not term or len(term) < 2:
            messagebox.showwarning("Search", "Enter at least 2 characters", parent=self)
            return

        try:
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            df = df.astype(str)

            mask = (
                df["CID"].str.lower().str.contains(term)
                | df["Name"].str.lower().str.contains(term)
                | df["Phone"].str.contains(term)
                | df["Address"].str.lower().str.contains(term)
            )

            filtered = df[mask]
            self.tree.delete(*self.tree.get_children())

            for idx, row in enumerate(filtered.itertuples(), start=0):
                tag = "oddrow" if idx % 2 else "evenrow"
                self.tree.insert(
                    "", "end",
                    values=(row._1, row.CID, row.Name, row.Phone, row.Address),
                    tags=(tag,)
                )

            self._focus_first_row()   # ← REQUIRED

        except Exception as e:
            messagebox.showerror("Search Error", f"Search failed: {e}", parent=self)
            self.load_data()


    def edit_customer(self):
        """Edit selected customer - FIXED BUTTON VISIBILITY"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Edit", "Select a customer to edit", parent=self)
            return

        if len(selected_items) > 1:
            messagebox.showwarning("Edit", "Select only one customer at a time", parent=self)
            return

        item = selected_items[0]
        values = self.tree.item(item)["values"]
        cid = str(values[1])
        old_name = str(values[2])
        old_phone = str(values[3])
        old_address = str(values[4])

        # ---------------- WINDOW ----------------
        edit_window = ctk.CTkToplevel(self)
        edit_window.title(f"Edit Customer - {cid}")
        edit_window.geometry("650x500")
        edit_window.transient(self)
        edit_window.grab_set()

        edit_window.update_idletasks()
        x = (edit_window.winfo_screenwidth() // 2) - 325
        y = (edit_window.winfo_screenheight() // 2) - 250
        edit_window.geometry(f"650x500+{x}+{y}")

        # Root container (VERTICAL)
        root = ctk.CTkFrame(edit_window)
        root.pack(fill="both", expand=True)

        # ---------------- HEADER ----------------
        header = ctk.CTkFrame(root)
        header.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(
            header,
            text="✏️ Edit Customer Details",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=self.colors["primary"]
        ).pack()

        # ---------------- FORM ----------------
        form_frame = ctk.CTkFrame(root)
        form_frame.pack(fill="both", expand=True, padx=20, pady=10)

        ctk.CTkLabel(form_frame, text="Name", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        ent_name = ctk.CTkEntry(form_frame, height=45)
        ent_name.insert(0, old_name)
        ent_name.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(form_frame, text="Phone", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        ent_phone = ctk.CTkEntry(form_frame, height=45)
        ent_phone.insert(0, old_phone)
        ent_phone.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(form_frame, text="Address", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        ent_address = ctk.CTkEntry(form_frame, height=45)
        ent_address.insert(0, old_address)
        ent_address.pack(fill="x")

        # ---------------- SAVE LOGIC ----------------
        def save_customer_changes():
            new_name = ent_name.get().strip()
            new_phone = ent_phone.get().strip()
            new_address = ent_address.get().strip()

            if not new_name:
                messagebox.showerror("Error", "Name is required", parent=edit_window)
                return
            if not new_phone.isdigit() or len(new_phone) != 10:
                messagebox.showerror("Error", "Enter valid 10-digit phone number", parent=edit_window)
                return
            if not new_address:
                messagebox.showerror("Error", "Address is required", parent=edit_window)
                return

            try:
                df = pd.read_excel(self.excel_file, engine="openpyxl")
                df["CID"] = df["CID"].astype(str)
                df["Phone"] = df["Phone"].astype(str)

                mask = df["CID"] == cid
                df.loc[mask, "Name"] = new_name
                df.loc[mask, "Phone"] = new_phone
                df.loc[mask, "Address"] = new_address
                df.to_excel(self.excel_file, index=False, engine="openpyxl")

                self.load_data()
                if self.on_customer_change:
                    self.on_customer_change()

                messagebox.showinfo("Success", "Customer updated successfully", parent=edit_window)
                edit_window.destroy()

            except Exception as e:
                messagebox.showerror("Save Error", str(e), parent=edit_window)

        # ---------------- BUTTON BAR (FIXED) ----------------
        button_bar = ctk.CTkFrame(root)
        button_bar.pack(fill="x", padx=20, pady=15)

        ctk.CTkButton(
            button_bar,
            text="💾 SAVE CHANGES",
            font=ctk.CTkFont(size=16, weight="bold"),
            width=220,
            height=50,
            fg_color="#059669",
            hover_color="#047857",
            command=save_customer_changes
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            button_bar,
            text="❌ CANCEL",
            font=ctk.CTkFont(size=16, weight="bold"),
            width=160,
            height=50,
            fg_color="#6b7280",
            hover_color="#4b5563",
            command=edit_window.destroy
        ).pack(side="left", padx=10)

    

            



    def delete_customers(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Delete", "Select at least one customer", parent=self)
            return
        
        cids = [
            self.tree.item(i)["values"][1]
            for i in selected_items
            if self.tree.item(i)["values"][1].startswith("C_")
        ]
        
        if not cids:
            messagebox.showwarning("Delete", "Select valid customer rows", parent=self)
            return
        
        confirm = messagebox.askyesno(
            "Delete", f"Delete {len(cids)} customer(s)?", parent=self
        )
        if not confirm:
            return
        
        try:
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            deleted_df = pd.read_excel(self.deleted_file, engine="openpyxl")

            to_delete = df[df["CID"].isin(cids)]
            
            # SAVE UNDO STATE BEFORE DELETING
            self.save_undo_state('delete', to_delete)
            
            deleted_df = pd.concat([deleted_df, to_delete], ignore_index=True)
            deleted_df["S.No"] = range(1, len(deleted_df) + 1)
            deleted_df.to_excel(self.deleted_file, index=False, engine="openpyxl")

            df = df[~df["CID"].isin(cids)]
            df["S.No"] = range(1, len(df) + 1)
            df.to_excel(self.excel_file, index=False, engine="openpyxl")

            self.load_data()
            self.lbl_cid.configure(text=self.get_next_cid())
            messagebox.showinfo("Delete", f"Deleted {len(cids)} customers", parent=self)
            
            if self.on_customer_change:
                self.on_customer_change()
                
        except Exception as e:
            messagebox.showerror("Error", f"Delete failed: {e}", parent=self)