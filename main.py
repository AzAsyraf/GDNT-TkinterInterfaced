import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import re
import os
from tkinter import ttk

# Main interface creation and event handling
def create_interface():
    # Handles file upload and processing
    def upload_and_process():
        file_path = filedialog.askopenfilename(
            filetypes=[("STEP or text files", "*.step *.stp *.txt"),
                       ("All files", "*.*")]
        )
        if file_path:
            filename_label.config(text=f"📁 File loaded: {os.path.basename(file_path)}")
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                    content = file.read()
                    # Get result as list of rows for table
                    result_lines = extract_tolerance_table(content)
                    show_table(result_lines)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")

    # Extracts tolerance values and datums from STEP/text file
    def extract_tolerance_values(text):
        lines = text.splitlines()
        # Build a dictionary mapping entity IDs to their lines
        line_dict = {
            re.match(r"(#\d+)\s*=", line).group(1): line.strip()
            for line in lines if re.match(r"(#\d+)\s*=", line)
        }

        # Regex patterns for tolerance, datum, and shape aspect entities
        tol_pattern = re.compile(
            r"(#\d+)\s*=\s*(CYLINDRICITY|FLATNESS|STRAIGHTNESS|ROUNDNESS)_TOLERANCE"
            r"\(\s*'([^']*)'\s*,\s*''\s*,\s*(#\d+)", re.IGNORECASE
        )
        datum_pattern = re.compile(
            r"#\d+\s*=\s*DATUM_FEATURE\(\s*'[^']*?\((\w)\)'", re.IGNORECASE
        )
        shape_aspect_pattern = re.compile(
            r"#\d+\s*=\s*SHAPE_ASPECT\('([^']*?)\((\w)?'?,.*?#(\d+)\)"
        )

        tol_results = []  # List of tolerance results
        datum_results = {}  # Mapping of datum letters to locations
        face_to_plane = {}  # Mapping of face IDs to surface types

        # Build datum letter to face ID and feature name mapping
        datum_letter_to_faceid = {}
        faceid_to_name = {}
        # Find all DATUM and SHAPE_ASPECT lines
        for line in lines:
            # Parse DATUM entity
            m = re.match(r"#\d+=DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", line)
            if m:
                feature, letter = m.groups()
                # Find corresponding SHAPE_ASPECT for this feature
                for sa_line in lines:
                    sa_m = re.match(r"#(\d+)=SHAPE_ASPECT\('([^']*)','',#\d+,\.T\.\);", sa_line)
                    if sa_m and feature in sa_m.group(2):
                        faceid = sa_m.group(1)
                        datum_letter_to_faceid[letter] = faceid
                        faceid_to_name[faceid] = feature
            # Parse SHAPE_ASPECT entity
            sa_m = re.match(r"#(\d+)=SHAPE_ASPECT\('([^']*)','',#\d+,\.T\.\);", line)
            if sa_m:
                faceid, feature = sa_m.groups()
                faceid_to_name[faceid] = feature

        # Map face IDs to shape descriptions (Plane1, Plane2, Boss1, etc.)
        for match in shape_aspect_pattern.finditer(text):
            shape_name, datum_letter, plane_id = match.groups()
            shape_name = shape_name.lower()
            location = ""
            if "plane1" in shape_name:
                location = "Plane1"
            elif "plane2" in shape_name:
                location = "Plane2"
            elif "boss1" in shape_name:
                location = "Boss1"
            elif "top" in shape_name:
                location = "top face"
            elif "bottom" in shape_name:
                location = "bottom face"
            elif "cylindrical" in shape_name or "side" in shape_name:
                location = "cylindrical side"
            face_to_plane[plane_id] = location
            if datum_letter:
                datum_results[datum_letter] = location

        # Extract tolerances and associate with datums and locations
        for tol_id, tol_type, tol_name, ref_id in tol_pattern.findall(text):
            definition = line_dict.get(ref_id, "")
            value_match = re.search(
                r"(?:LENGTH_MEASURE|VALUE_REPRESENTATION_ITEM)\s*\(\s*([\d.]+)", definition
            )
            value = f"±{value_match.group(1)}" if value_match else "N/A"
            label = "Circularity" if tol_type.upper() == "ROUNDNESS" else tol_type.capitalize()

            # Find the last # reference in the tolerance line for datum mapping
            tol_line = next((l for l in lines if tol_id in l), "")
            ref_ids = re.findall(r"#(\d+)", tol_line)
            datum_ref_id = ref_ids[-1] if ref_ids else ref_id

            # Map datum_ref_id to datum letter and location
            datum_letter = ""
            location = ""
            for letter, faceid in datum_letter_to_faceid.items():
                if faceid == datum_ref_id:
                    datum_letter = letter
                    location = faceid_to_name.get(faceid, face_to_plane.get(faceid, ""))
                    break
            if not datum_letter:
                # Fallback to previous logic if not found
                tol_name_lower = tol_name.lower()
                for d_letter in datum_results:
                    if f"({d_letter.lower()})" in tol_name_lower:
                        datum_letter = d_letter
                        break
                if datum_letter and datum_letter in datum_letter_to_faceid:
                    faceid = datum_letter_to_faceid[datum_letter]
                    location = faceid_to_name.get(faceid, face_to_plane.get(faceid, ""))
                else:
                    for key, val in face_to_plane.items():
                        if key in ref_id or key in tol_name:
                            location = val
                            break

            tol_results.append((label, value, datum_letter, location))

        # GD&T symbol mapping for output
        gdnt_symbols = {
            "Straightness": "─",
            "Flatness": "☐",
            "Circularity": "○",
            "Cylindricity": "⌀"
        }

        # Helper to map feature name to surface type (for Location column)
        def get_surface_type(feature_name):
            fname = feature_name.lower()
            if "plane1" in fname or "top" in fname:
                return "top face"
            elif "plane2" in fname or "bottom" in fname:
                return "bottom face"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "cylindrical side"
            else:
                return feature_name

        # Helper to map feature name and type to Surface (for Surface column)
        def get_likely_location(label, feature_name):
            fname = feature_name.lower()
            if "plane1" in fname or "top" in fname:
                return "top face (facing +Z)"
            elif "plane2" in fname or "bottom" in fname:
                return "bottom face (facing -Z)"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "curved side of the cylinder"
            elif "face" in fname:
                return "planar face"
            else:
                return feature_name

        # Parse DIRECTION and AXIS2_PLACEMENT_3D for axis orientation
        direction_map = {}  # Maps DIRECTION entity IDs to vectors
        axis_map = {}       # Maps AXIS2_PLACEMENT_3D IDs to direction IDs
        plane_map = {}      # Maps PLANE IDs to AXIS2_PLACEMENT_3D IDs
        for line in lines:
            # Parse DIRECTION entity
            dir_match = re.match(r"#(\d+)=DIRECTION\('[^']*',\(([^)]*)\)\);", line)
            if dir_match:
                dir_id, vec = dir_match.groups()
                vec = tuple(float(x) for x in vec.split(","))
                direction_map[dir_id] = vec
            # Parse AXIS2_PLACEMENT_3D entity
            axis_match = re.match(r"#(\d+)=AXIS2_PLACEMENT_3D\('[^']*',#(\d+),#(\d+),#(\d+)\);", line)
            if axis_match:
                axis_id, pt_id, dir1_id, dir2_id = axis_match.groups()
                axis_map[axis_id] = (dir1_id, dir2_id)
            # Parse PLANE entity
            plane_match = re.match(r"#(\d+)=PLANE\('[^']*',#(\d+)\);", line)
            if plane_match:
                plane_id, axis_id = plane_match.groups()
                plane_map[plane_id] = axis_id

        # Helper to get axis orientation from plane_id
        def get_axis_orientation(plane_id):
            axis_id = plane_map.get(plane_id)
            if not axis_id:
                return "unknown axis"
            dir_ids = axis_map.get(axis_id)
            if not dir_ids:
                return "unknown axis"
            # Usually the first direction is the normal
            dir_vec = direction_map.get(dir_ids[0])
            if not dir_vec:
                return "unknown axis"
            # Find closest axis
            axis_labels = ["X", "Y", "Z"]
            axis_vectors = [(1,0,0), (0,1,0), (0,0,1)]
            dot_products = [abs(sum(a*b for a,b in zip(dir_vec, axis))) for axis in axis_vectors]
            max_idx = dot_products.index(max(dot_products))
            sign = '+' if dir_vec[max_idx] >= 0 else '-'
            return f"facing {sign}{axis_labels[max_idx]}"

        # Build output text for results
        output = f"{'Type':<18}{'Value':<10}{'Datum':<9}{'Location':<18}{'Surface':<25}\n" + "-" * 80 + "\n"
        for label, value, datum, loc in tol_results:
            symbol = gdnt_symbols.get(label, "")
            type_with_symbol = f"{symbol} {label}" if symbol else label
            location_str = get_surface_type(loc)
            likely_location = get_likely_location(label, loc)
            # If planar face, try to get axis orientation
            axis_info = ""
            for pid in plane_map:
                if location_str.startswith("top face") or location_str.startswith("bottom face"):
                    axis_info = get_axis_orientation(pid)
                    break
            if axis_info and axis_info != "unknown axis":
                likely_location = f"{location_str} ({axis_info})"
            output += f"{type_with_symbol:<18}{value:<10}{datum:<9}{location_str:<18}{likely_location:<25}\n\n"
        # Output datums with their mapped locations
        for d_letter in datum_letter_to_faceid:
            faceid = datum_letter_to_faceid[d_letter]
            feature_name = faceid_to_name.get(faceid, "")
            location_str = get_surface_type(feature_name)
            likely_location = get_likely_location('Datum', feature_name)
            axis_info = ""
            for pid in plane_map:
                if location_str.startswith("top face") or location_str.startswith("bottom face"):
                    axis_info = get_axis_orientation(pid)
                    break
            if axis_info and axis_info != "unknown axis":
                likely_location = f"{location_str} ({axis_info})"
            output += f"{'Datum':<18}{d_letter:<10}{d_letter:<9}{location_str:<18}{likely_location:<25}\n\n"
        if not tol_results and not datum_results:
            return "⚠️ No tolerance or datum data found."
        return output

    # New function to extract table rows for Treeview
    def extract_tolerance_table(text):
        lines = text.splitlines()
        # Build a dictionary mapping entity IDs to their lines
        line_dict = {
            re.match(r"(#\d+)\s*=", line).group(1): line.strip()
            for line in lines if re.match(r"(#\d+)\s*=", line)
        }

        # Regex patterns for tolerance, datum, and shape aspect entities
        tol_pattern = re.compile(
            r"(#\d+)\s*=\s*(CYLINDRICITY|FLATNESS|STRAIGHTNESS|ROUNDNESS)_TOLERANCE"
            r"\(\s*'([^']*)'\s*,\s*''\s*,\s*(#\d+)", re.IGNORECASE
        )
        datum_pattern = re.compile(
            r"#\d+\s*=\s*DATUM_FEATURE\(\s*'[^']*?\((\w)\)'", re.IGNORECASE
        )
        shape_aspect_pattern = re.compile(
            r"#\d+\s*=\s*SHAPE_ASPECT\('([^']*?)\((\w)?'?,.*?#(\d+)\)"
        )

        tol_results = []  # List of tolerance results
        datum_results = {}  # Mapping of datum letters to locations
        face_to_plane = {}  # Mapping of face IDs to surface types

        # Build datum letter to face ID and feature name mapping
        datum_letter_to_faceid = {}
        faceid_to_name = {}
        # Find all DATUM and SHAPE_ASPECT lines
        for line in lines:
            # Parse DATUM entity
            m = re.match(r"#\d+=DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", line)
            if m:
                feature, letter = m.groups()
                # Find corresponding SHAPE_ASPECT for this feature
                for sa_line in lines:
                    sa_m = re.match(r"#(\d+)=SHAPE_ASPECT\('([^']*)','',#\d+,\.T\.\);", sa_line)
                    if sa_m and feature in sa_m.group(2):
                        faceid = sa_m.group(1)
                        datum_letter_to_faceid[letter] = faceid
                        faceid_to_name[faceid] = feature
            # Parse SHAPE_ASPECT entity
            sa_m = re.match(r"#(\d+)=SHAPE_ASPECT\('([^']*)','',#\d+,\.T\.\);", line)
            if sa_m:
                faceid, feature = sa_m.groups()
                faceid_to_name[faceid] = feature

        # Map face IDs to shape descriptions (Plane1, Plane2, Boss1, etc.)
        for match in shape_aspect_pattern.finditer(text):
            shape_name, datum_letter, plane_id = match.groups()
            shape_name = shape_name.lower()
            location = ""
            if "plane1" in shape_name:
                location = "Plane1"
            elif "plane2" in shape_name:
                location = "Plane2"
            elif "boss1" in shape_name:
                location = "Boss1"
            elif "top" in shape_name:
                location = "top face"
            elif "bottom" in shape_name:
                location = "bottom face"
            elif "cylindrical" in shape_name or "side" in shape_name:
                location = "cylindrical side"
            face_to_plane[plane_id] = location
            if datum_letter:
                datum_results[datum_letter] = location

        # Extract tolerances and associate with datums and locations
        for tol_id, tol_type, tol_name, ref_id in tol_pattern.findall(text):
            definition = line_dict.get(ref_id, "")
            value_match = re.search(
                r"(?:LENGTH_MEASURE|VALUE_REPRESENTATION_ITEM)\s*\(\s*([\d.]+)", definition
            )
            value = f"±{value_match.group(1)}" if value_match else "N/A"
            label = "Circularity" if tol_type.upper() == "ROUNDNESS" else tol_type.capitalize()

            # Find the last # reference in the tolerance line for datum mapping
            tol_line = next((l for l in lines if tol_id in l), "")
            ref_ids = re.findall(r"#(\d+)", tol_line)
            datum_ref_id = ref_ids[-1] if ref_ids else ref_id

            # Map datum_ref_id to datum letter and location
            datum_letter = ""
            location = ""
            for letter, faceid in datum_letter_to_faceid.items():
                if faceid == datum_ref_id:
                    datum_letter = letter
                    location = faceid_to_name.get(faceid, face_to_plane.get(faceid, ""))
                    break
            if not datum_letter:
                # Fallback to previous logic if not found
                tol_name_lower = tol_name.lower()
                for d_letter in datum_results:
                    if f"({d_letter.lower()})" in tol_name_lower:
                        datum_letter = d_letter
                        break
                if datum_letter and datum_letter in datum_letter_to_faceid:
                    faceid = datum_letter_to_faceid[datum_letter]
                    location = faceid_to_name.get(faceid, face_to_plane.get(faceid, ""))
                else:
                    for key, val in face_to_plane.items():
                        if key in ref_id or key in tol_name:
                            location = val
                            break

            tol_results.append((label, value, datum_letter, location))

        # GD&T symbol mapping for output
        gdnt_symbols = {
            "Straightness": "─",
            "Flatness": "☐",
            "Circularity": "○",
            "Cylindricity": "⌀"
        }

        # Helper to map feature name to surface type (for Location column)
        def get_surface_type(feature_name):
            fname = feature_name.lower()
            if "plane1" in fname or "top" in fname:
                return "top face"
            elif "plane2" in fname or "bottom" in fname:
                return "bottom face"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "cylindrical side"
            else:
                return feature_name

        # Helper to map feature name and type to Surface (for Surface column)
        def get_likely_location(label, feature_name):
            fname = feature_name.lower()
            if "plane1" in fname or "top" in fname:
                return "top face (facing +Z)"
            elif "plane2" in fname or "bottom" in fname:
                return "bottom face (facing -Z)"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "curved side of the cylinder"
            elif "face" in fname:
                return "planar face"
            else:
                return feature_name

        # Parse DIRECTION and AXIS2_PLACEMENT_3D for axis orientation
        direction_map = {}  # Maps DIRECTION entity IDs to vectors
        axis_map = {}       # Maps AXIS2_PLACEMENT_3D IDs to direction IDs
        plane_map = {}      # Maps PLANE IDs to AXIS2_PLACEMENT_3D IDs
        for line in lines:
            # Parse DIRECTION entity
            dir_match = re.match(r"#(\d+)=DIRECTION\('[^']*',\(([^)]*)\)\);", line)
            if dir_match:
                dir_id, vec = dir_match.groups()
                vec = tuple(float(x) for x in vec.split(","))
                direction_map[dir_id] = vec
            # Parse AXIS2_PLACEMENT_3D entity
            axis_match = re.match(r"#(\d+)=AXIS2_PLACEMENT_3D\('[^']*',#(\d+),#(\d+),#(\d+)\);", line)
            if axis_match:
                axis_id, pt_id, dir1_id, dir2_id = axis_match.groups()
                axis_map[axis_id] = (dir1_id, dir2_id)
            # Parse PLANE entity
            plane_match = re.match(r"#(\d+)=PLANE\('[^']*',#(\d+)\);", line)
            if plane_match:
                plane_id, axis_id = plane_match.groups()
                plane_map[plane_id] = axis_id

        # Helper to get axis orientation from plane_id
        def get_axis_orientation(plane_id):
            axis_id = plane_map.get(plane_id)
            if not axis_id:
                return "unknown axis"
            dir_ids = axis_map.get(axis_id)
            if not dir_ids:
                return "unknown axis"
            dir_vec = direction_map.get(dir_ids[0])
            if not dir_vec:
                return "unknown axis"
            axis_labels = ["X", "Y", "Z"]
            axis_vectors = [(1,0,0), (0,1,0), (0,0,1)]
            dot_products = [abs(sum(a*b for a,b in zip(dir_vec, axis))) for axis in axis_vectors]
            max_idx = dot_products.index(max(dot_products))
            sign = '+' if dir_vec[max_idx] >= 0 else '-'
            return f"facing {sign}{axis_labels[max_idx]}"

        # Build output text for results
        table_rows = []
        for label, value, datum, loc in tol_results:
            symbol = gdnt_symbols.get(label, "")
            type_with_symbol = f"{symbol} {label}" if symbol else label
            location_str = get_surface_type(loc)
            surface = get_likely_location(label, loc)
            axis_info = ""
            for pid in plane_map:
                if location_str.startswith("top face") or location_str.startswith("bottom face"):
                    axis_info = get_axis_orientation(pid)
                    break
            if axis_info and axis_info != "unknown axis":
                surface = f"{location_str} ({axis_info})"
            table_rows.append((type_with_symbol, value, datum, location_str, surface))
        for d_letter in datum_letter_to_faceid:
            faceid = datum_letter_to_faceid[d_letter]
            feature_name = faceid_to_name.get(faceid, "")
            location_str = get_surface_type(feature_name)
            surface = get_likely_location('Datum', feature_name)
            axis_info = ""
            for pid in plane_map:
                if location_str.startswith("top face") or location_str.startswith("bottom face"):
                    axis_info = get_axis_orientation(pid)
                    break
            if axis_info and axis_info != "unknown axis":
                surface = f"{location_str} ({axis_info})"
            table_rows.append(("Datum", d_letter, d_letter, location_str, surface))
        return table_rows

    # Displays the results in a ttk.Treeview table
    def show_table(result_lines):
        # Destroy previous table if exists
        global output_table
        if output_table:
            output_table.destroy()
        columns = ("Type", "Value", "Datum", "Location", "Surface")
        output_table = ttk.Treeview(root, columns=columns, show="headings", height=20)
        for col in columns:
            output_table.heading(col, text=col)
            output_table.column(col, anchor="center", width=150)
        # Configure tags for alternating row colors
        output_table.tag_configure('even', background='#f5f5f5')
        output_table.tag_configure('odd', background='#e9ecef')
        for idx, line in enumerate(result_lines):
            tag = 'even' if idx % 2 == 0 else 'odd'
            output_table.insert("", "end", values=line, tags=(tag,))
        output_table.pack(expand=True, fill='both', padx=20, pady=10)

    # Handles saving results to file
    def save_results():
        global output_table
        rows = [output_table.item(row)['values'] for row in output_table.get_children()]
        if not rows:
            messagebox.showwarning("Warning", "Nothing to save.")
            return
        file_format = format_var.get()
        filetypes = [("Text files", "*.txt")] if file_format == ".txt" else \
                    [("CSV files", "*.csv")] if file_format == ".csv" else \
                    [("Excel files", "*.xlsx")]
        file_path = filedialog.asksaveasfilename(defaultextension=file_format,
                                                 filetypes=filetypes)
        if not file_path:
            return
        try:
            headers = ["Type", "Value", "Datum", "Location", "Surface"]
            if file_format == ".txt":
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write("\t".join(headers) + "\n")
                    for row in rows:
                        file.write("\t".join(str(cell) for cell in row) + "\n")
            elif file_format == ".csv":
                import csv
                with open(file_path, "w", newline='', encoding="utf-8") as file:
                    writer = csv.writer(file)
                    writer.writerow(headers)
                    writer.writerows(rows)
            elif file_format == ".xlsx":
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "GD&T Tolerances"
                ws.append(headers)
                for row in rows:
                    ws.append(row)
                wb.save(file_path)
            messagebox.showinfo("Saved", f"Results saved as {file_format.upper()}!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    # Clears the output text and filename label
    def clear_output():
        global output_table
        if output_table:
            for row in output_table.get_children():
                output_table.delete(row)
        filename_label.config(text="")

    # Toggles between dark and bright mode
    def toggle_theme():
        nonlocal dark_mode
        dark_mode = not dark_mode
        bg_color = "#2c2f33" if dark_mode else "#eef3f7"
        fg_color = "#f0f0f0" if dark_mode else "#003366"
        root.configure(bg=bg_color)
        top_frame.configure(bg=bg_color)
        mid_frame.configure(bg=bg_color)
        title_label.configure(bg=bg_color, fg=fg_color)
        filename_label.configure(bg=bg_color, fg="#bbbbbb" if dark_mode else "black")
        output_text.configure(bg="#1e1e1e" if dark_mode else "white",
                              fg="#f0f0f0" if dark_mode else "#333333",
                              insertbackground="white" if dark_mode else "black")
        theme_button.configure(
            text="☀️ Switch to Bright Mode" if dark_mode else "🌙 Switch to Dark Mode",
            bg="#7289da" if dark_mode else "#dddddd",
            fg="white" if dark_mode else "black")

    # GUI Setup
    root = tk.Tk()
    root.title("GD&T Tolerance")
    root.geometry("850x650")
    root.configure(bg="#eef3f7")
    dark_mode = False

    top_frame = tk.Frame(root, bg="#eef3f7")
    top_frame.pack(fill=tk.X, padx=20, pady=10)
    top_frame.grid_columnconfigure(0, weight=1)
    top_frame.grid_columnconfigure(1, weight=3)
    top_frame.grid_columnconfigure(2, weight=1)

    logo_frame = tk.Frame(top_frame, bg="#ffffff", borderwidth=1, relief="solid")
    logo_frame.grid(row=0, column=0, sticky="w", padx=10)
    try:
        logo_img = Image.open("./assets/unisza.png").resize((150, 150))
        logo_photo = ImageTk.PhotoImage(logo_img)
        tk.Label(logo_frame, image=logo_photo, bg="#ffffff").pack()
    except: pass

    title_label = tk.Label(top_frame, text="GD&T Tolerance Value Extractor",
                            font=("Helvetica", 20, "bold"), bg="#eef3f7", fg="#003366")
    title_label.grid(row=0, column=1, padx=10, pady=10)

    right_logo_frame = tk.Frame(top_frame, bg="#ffffff", borderwidth=1, relief="solid")
    right_logo_frame.grid(row=0, column=2, sticky="e", padx=10)
    try:
        right_img = Image.open("./assets/frit.png").resize((150, 150))
        right_photo = ImageTk.PhotoImage(right_img)
        tk.Label(right_logo_frame, image=right_photo, bg="#ffffff").pack()
    except: pass

    mid_frame = tk.Frame(root, bg="#eef3f7")
    mid_frame.pack(pady=5)

    tk.Button(mid_frame, text="📤 Upload STEP File", command=upload_and_process,
              font=("Arial", 11, "bold"), bg="#4CAF50", fg="white").grid(row=0, column=0, padx=5)

    tk.Button(mid_frame, text="💾 Save Result", command=save_results,
              font=("Arial", 11), bg="#4CAF50", fg="white").grid(row=0, column=1, padx=5)

    tk.Button(mid_frame, text="🗑️ Clear Output", command=clear_output,
              font=("Arial", 11), bg="#ffcc00", fg="black").grid(row=0, column=2, padx=5)

    theme_button = tk.Button(mid_frame, text="🌙 Switch to Dark Mode", command=toggle_theme,
                             font=("Arial", 10), bg="#dddddd", fg="black")
    theme_button.grid(row=0, column=3, padx=5)

    format_var = tk.StringVar(value=".txt")
    tk.OptionMenu(mid_frame, format_var, ".txt", ".csv", ".xlsx").grid(row=0, column=4, padx=5)

    filename_label = tk.Label(root, text="", font=("Arial", 10), bg="#eef3f7", fg="black")
    filename_label.pack(pady=(5, 0))

    # Create empty table on startup
    global output_table
    columns = ("Type", "Value", "Datum", "Location", "Surface")
    output_table = ttk.Treeview(root, columns=columns, show="headings", height=20)
    for col in columns:
        output_table.heading(col, text=col)
        output_table.column(col, anchor="center", width=150)
    output_table.tag_configure('even', background='#f5f5f5')
    output_table.tag_configure('odd', background='#e9ecef')
    output_table.pack(expand=True, fill='both', padx=20, pady=10)

    # --- Check Datum A and B orientation from STEP data ---
    step_text = '''
#137=PLANE('',#136);
#136=AXIS2_PLACEMENT_3D('',#133,#134,#135);
#133=CARTESIAN_POINT('',(0.,0.,0.));
#134=DIRECTION('',(0.,0.,1.));
#135=DIRECTION('',(1.,0.,0.));
#231=PLANE('',#230);
#230=AXIS2_PLACEMENT_3D('',#227,#228,#229);
#227=CARTESIAN_POINT('',(10.,0.,0.));
#228=DIRECTION('',(1.,0.,0.));
#229=DIRECTION('',(0.,0.,-1.));
'''
    lines = step_text.splitlines()
    plane_id_to_axis = {}
    axis_id_to_dir = {}
    dir_id_to_vec = {}
    for line in lines:
        # Parse PLANE entity
        plane_match = re.match(r"#(\d+)=PLANE\('[^']*',#(\d+)\);", line)
        if plane_match:
            pid, aid = plane_match.groups()
            plane_id_to_axis[pid] = aid
        # Parse AXIS2_PLACEMENT_3D entity
        axis_match = re.match(r"#(\d+)=AXIS2_PLACEMENT_3D\('[^']*',#(\d+),#(\d+),#(\d+)\);", line)
        if axis_match:
            aid, ptid, dir1id, dir2id = axis_match.groups()
            axis_id_to_dir[aid] = (dir1id, dir2id)
        # Parse DIRECTION entity
        dir_match = re.match(r"#(\d+)=DIRECTION\('[^']*',\(([^)]*)\)\);", line)
        if dir_match:
            did, vec = dir_match.groups()
            vec = tuple(float(x) for x in vec.split(","))
            dir_id_to_vec[did] = vec
    # Helper to get axis orientation from plane_id
    def get_axis_orientation(plane_id):
        axis_id = plane_id_to_axis.get(plane_id)
        if not axis_id:
            return "unknown axis"
        dir_ids = axis_id_to_dir.get(axis_id)
        if not dir_ids:
            return "unknown axis"
        dir_vec = dir_id_to_vec.get(dir_ids[0])
        if not dir_vec:
            return "unknown axis"
        axis_labels = ["X", "Y", "Z"]
        axis_vectors = [(1,0,0), (0,1,0), (0,0,1)]
        dot_products = [abs(sum(a*b for a,b in zip(dir_vec, axis))) for axis in axis_vectors]
        max_idx = dot_products.index(max(dot_products))
        sign = '+' if dir_vec[max_idx] >= 0 else '-'
        return f"facing {sign}{axis_labels[max_idx]}"
    # Example usage (can be used for debugging or further logic)
    datum_a_orientation = get_axis_orientation('137')
    datum_b_orientation = get_axis_orientation('231')

    root.mainloop()

create_interface()
