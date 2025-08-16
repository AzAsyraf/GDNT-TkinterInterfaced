import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import re
import os
from tkinter import ttk

# Main interface creation and event handling
def create_interface():
    # Enhanced function to determine surface position for datum plane detection
    def determine_plane_position(feature_name, datum_letter=None, plane_data=None):
        """
        Enhanced function to determine if a plane datum is on top face or bottom face
        """
        if not feature_name:
            return "surface"
            
        fname = str(feature_name).lower()
        
        # Direct keyword detection for position
        if any(keyword in fname for keyword in ["top", "upper", "above"]):
            return "top face"
        elif any(keyword in fname for keyword in ["bottom", "lower", "below", "base"]):
            return "bottom face"
        
        # Enhanced datum letter logic - based on common CAD conventions
        if datum_letter:
            datum_upper = datum_letter.upper()
            # In many CAD systems, datum A is often the primary reference (base/bottom)
            # and subsequent datums (B, C, D, etc.) can be on different surfaces
            if datum_upper == 'A':
                # Check if it's explicitly a top surface first
                if any(keyword in fname for keyword in ["top", "upper"]):
                    return "top face"
                else:
                    return "bottom face"  # A is typically base/bottom
            elif datum_upper in ['B', 'C']:
                return "cylindrical side"  # B and C often on cylindrical surfaces
            elif datum_upper == 'D':
                # D is often the secondary planar surface (opposite of A)
                return "top face"
            elif datum_upper in ['E', 'F', 'G', 'H']:
                # Additional datums might alternate
                return "top face" if datum_upper in ['E', 'G'] else "bottom face"
        
        # Plane numbering logic (less reliable)
        elif "plane1" in fname:
            return "bottom face"  # Often plane1 is the base
        elif "plane2" in fname:
            return "top face"   # Often plane2 is the top
        
        # Check for numerical patterns that might indicate position
        plane_numbers = re.findall(r'plane(\d+)', fname)
        if plane_numbers:
            plane_num = int(plane_numbers[0])
            if plane_num == 1:
                return "bottom face"  # Reversed assumption
            elif plane_num >= 2:
                return "top face"
        
        # If it contains "plane" but no clear position indicator, try to infer
        if "plane" in fname:
            # Check if there are any Z-axis indicators or position clues
            if any(indicator in fname for indicator in ["+z", "positive", "high"]):
                return "top face"
            elif any(indicator in fname for indicator in ["-z", "negative", "low"]):
                return "bottom face"
            else:
                # Default assumption: return as generic plane
                return "planar surface"
        
        return str(feature_name)

    # Handles file upload and processing
    def upload_and_process():
        file_path = filedialog.askopenfilename(
            filetypes=[("STEP or text files", "*.step *.stp *.txt"),
                       ("All files", "*.*")]
        )
        if file_path:
            filename_label.config(
                text=f"📁 File loaded: {os.path.basename(file_path)}")
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                    content = file.read()
                    # Get result as list of rows for table
                    result_lines = extract_tolerance_table(content)
                    show_table(result_lines)
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to read file:\n{str(e)}")

    # Extracts tolerance values and datums from STEP/text file
    def extract_tolerance_values(text):
        lines = text.splitlines()
        # Build a dictionary mapping entity IDs to their lines
        line_dict = {
            re.match(r"(#\d+)\s*=", line).group(1): line.strip()
            for line in lines if re.match(r"(#\d+)\s*=", line)
        }

        # Enhanced regex patterns for tolerance, datum, and shape aspect entities
        tol_pattern = re.compile(
            r"(#\d+)\s*=\s*(CYLINDRICITY|FLATNESS|STRAIGHTNESS|ROUNDNESS)_TOLERANCE"
            r"\(\s*'([^']*)'\s*,\s*''\s*,\s*(#\d+)", re.IGNORECASE
        )
        
        # Updated datum pattern to handle the new format
        datum_pattern = re.compile(
            r"#\d+\s*=\s*DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", re.IGNORECASE
        )
        
        shape_aspect_pattern = re.compile(
            r"#\d+\s*=\s*SHAPE_ASPECT\('([^']*?)\((\w)?'?,.*?#(\d+)\)"
        )

        tol_results = []  # List of tolerance results
        datum_results = {}  # Mapping of datum letters to locations
        face_to_plane = {}  # Mapping of face IDs to surface types

        # Enhanced datum extraction - Build datum letter to feature name mapping
        datum_letter_to_feature = {}
        
        # Find all DATUM lines with the enhanced pattern
        for line in lines:
            # Updated regex to match your file format: #522=DATUM('Datum29@Boss1(A)',$,#23,.F.,'A');
            datum_match = re.match(
                r"#\d+=DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", line)
            if datum_match:
                feature_name, datum_letter = datum_match.groups()
                datum_letter_to_feature[datum_letter] = feature_name
                
                # Extract the geometric feature type from the feature name
                # Example: 'Datum29@Boss1(A)' -> 'Boss1', 'Datum28@Plane1(D)' -> 'Plane1'
                feature_match = re.search(r'@([^(]+)', feature_name)
                if feature_match:
                    geometric_feature = feature_match.group(1).lower()
                    if 'boss' in geometric_feature:
                        datum_results[datum_letter] = "cylindrical side"
                    elif 'plane1' in geometric_feature:
                        datum_results[datum_letter] = "bottom face"
                    elif 'plane2' in geometric_feature:
                        datum_results[datum_letter] = "top face"
                    elif 'plane' in geometric_feature:
                        datum_results[datum_letter] = determine_plane_position(geometric_feature, datum_letter)
                    else:
                        datum_results[datum_letter] = geometric_feature
                else:
                    # Fallback if no @ pattern found
                    datum_results[datum_letter] = determine_plane_position(feature_name, datum_letter)

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
            elif "torus" in shape_name:
                location = "torus side"
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
            value = f"{value_match.group(1)}" if value_match else "N/A"
            label = "Circularity" if tol_type.upper() == "ROUNDNESS" else tol_type.capitalize()

            # Enhanced datum letter and location detection
            datum_letter = ""
            location = ""
            
            # Method 1: Extract datum letter from tolerance name (e.g., "tolerance(A)")
            tol_name_lower = tol_name.lower()
            datum_match = re.search(r'\(([A-Z])\)', tol_name)
            if datum_match:
                datum_letter = datum_match.group(1)
                location = datum_results.get(datum_letter, "")
            
            # Method 2: Try to find datum letter from tolerance name with lowercase
            if not datum_letter:
                for d_letter in datum_results:
                    if f"({d_letter.lower()})" in tol_name_lower or f"({d_letter.upper()})" in tol_name:
                        datum_letter = d_letter
                        location = datum_results[d_letter]
                        break
            
            # Method 3: Check if tolerance name contains feature references
            if not location:
                # Look for Boss1, Plane1, Plane2 in tolerance name
                if "boss1" in tol_name_lower:
                    location = "cylindrical side"
                elif "plane1" in tol_name_lower:
                    location = "bottom face"
                elif "plane2" in tol_name_lower:
                    location = "top face"
                elif "plane" in tol_name_lower:
                    # Generic plane - try to determine from context
                    location = "planar surface"
            
            # Method 4: Infer location based on tolerance type if still not found
            if not location:
                if label.lower() in ["cylindricity", "circularity"]:
                    location = "cylindrical side"
                elif label.lower() in ["flatness"]:
                    location = "planar surface"
                elif label.lower() in ["straightness"]:
                    location = "surface"

            tol_results.append((label, value, datum_letter, location))

        # GD&T symbol mapping for output
        gdnt_symbols = {
            "Straightness": "─",
            "Flatness": "☐",
            "Circularity": "○",
            "Cylindricity": "⌀"
        }

        # Helper to clean and parse feature names
        def clean_feature_name(feature_name):
            if not feature_name:
                return ""
            # Remove common STEP file artifacts and extract meaningful parts
            fname = str(feature_name).lower()
            # Remove datum references like "datum10@" or similar patterns
            fname = re.sub(r'datum\d+@', '', fname)
            # Extract geometric feature types
            if "torus" in fname:
                return "torus"
            elif "plane" in fname:
                return "plane"
            elif "boss" in fname:
                return "boss"
            elif "cylinder" in fname:
                return "cylinder"
            elif "cone" in fname:
                return "cone"
            else:
                # Keep original if no pattern matches
                return feature_name

        # Enhanced helper to map feature name to surface type (for Location column)
        def get_surface_type(feature_name):
            cleaned_name = clean_feature_name(feature_name)
            fname = cleaned_name.lower()
            
            if "torus" in fname:
                return "torus side"
            elif "plane" in str(feature_name).lower():
                # Enhanced plane detection using the new function
                return determine_plane_position(feature_name)
            elif "cone" in fname or "conical" in fname:
                return "conical side of the part"
            elif "boss" in fname or "cylindrical" in fname or "side" in fname or "cylinder" in fname:
                return "cylindrical side"
            else:
                return "surface"

        # Enhanced helper to map feature name and type to Surface (for Surface column)
        def get_likely_location(label, feature_name):
            cleaned_name = clean_feature_name(feature_name)
            fname = cleaned_name.lower()
            
            if "torus" in fname:
                return "torus side"
            elif "plane" in str(feature_name).lower():
                # Enhanced plane detection using the new function
                return determine_plane_position(feature_name)
            elif "cone" in fname or "conical" in fname:
                return "conical side of the part"
            elif "boss" in fname or "cylindrical" in fname or "side" in fname or "cylinder" in fname:
                return "curved side of the cylinder"
            elif "face" in fname:
                return "planar face"
            else:
                return "surface"

        # Build output text for results
        output = f"{'Type':<18}{'Value':<10}{'Datum':<9}{'Location':<18}{'Surface':<25}\n" + "-" * 80 + "\n"
        for label, value, datum, loc in tol_results:
            symbol = gdnt_symbols.get(label, "")
            type_with_symbol = f"{symbol} {label}" if symbol else label
            
            # Create location string with datum reference
            if datum:
                location_str = f"at datum {datum}"
            else:
                location_str = loc if loc else "surface"
            
            # Map location to surface description
            if loc == "cylindrical side":
                likely_location = "curved side of the cylinder"
            elif loc == "bottom face":
                likely_location = "bottom face"
            elif loc == "top face":
                likely_location = "top face"
            elif loc == "planar surface":
                likely_location = "planar face"
            else:
                likely_location = loc if loc else "surface"
                
            output += f"{type_with_symbol:<18}{value:<10}{datum:<9}{location_str:<18}{likely_location:<25}\n\n"
        
        # Enhanced output datums with their mapped locations
        for d_letter, feature_name in datum_letter_to_feature.items():
            location_str = datum_results.get(d_letter, "surface")
            likely_location = location_str
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

        # Enhanced regex patterns for tolerance, datum, and shape aspect entities
        tol_pattern = re.compile(
            r"(#\d+)\s*=\s*(CYLINDRICITY|FLATNESS|STRAIGHTNESS|ROUNDNESS)_TOLERANCE"
            r"\(\s*'([^']*)'\s*,\s*''\s*,\s*(#\d+)", re.IGNORECASE
        )
        
        # Updated datum pattern to handle the new format
        datum_pattern = re.compile(
            r"#\d+\s*=\s*DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", re.IGNORECASE
        )
        
        shape_aspect_pattern = re.compile(
            r"#\d+\s*=\s*SHAPE_ASPECT\('([^']*?)\((\w)?'?,.*?#(\d+)\)"
        )

        tol_results = []  # List of tolerance results
        datum_results = {}  # Mapping of datum letters to locations
        face_to_plane = {}  # Mapping of face IDs to surface types

        # Enhanced datum extraction - Build datum letter to feature name mapping
        datum_letter_to_feature = {}
        
        # Find all DATUM lines with the enhanced pattern
        for line in lines:
            # Updated regex to match your file format: #522=DATUM('Datum29@Boss1(A)',$,#23,.F.,'A');
            datum_match = re.match(
                r"#\d+=DATUM\('([^']*)',\$,#\d+,\.F\.,'([A-Z])'\);", line)
            if datum_match:
                feature_name, datum_letter = datum_match.groups()
                datum_letter_to_feature[datum_letter] = feature_name
                
                # Extract the geometric feature type from the feature name
                # Example: 'Datum29@Boss1(A)' -> 'Boss1', 'Datum28@Plane1(D)' -> 'Plane1'
                feature_match = re.search(r'@([^(]+)', feature_name)
                if feature_match:
                    geometric_feature = feature_match.group(1).lower()
                    if 'boss' in geometric_feature:
                        datum_results[datum_letter] = "cylindrical side"
                    elif 'plane1' in geometric_feature:
                        datum_results[datum_letter] = "bottom face"
                    elif 'plane2' in geometric_feature:
                        datum_results[datum_letter] = "top face"
                    elif 'plane' in geometric_feature:
                        datum_results[datum_letter] = determine_plane_position(geometric_feature, datum_letter)
                    else:
                        datum_results[datum_letter] = geometric_feature
                else:
                    # Fallback if no @ pattern found
                    datum_results[datum_letter] = determine_plane_position(feature_name, datum_letter)

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
            elif "torus" in shape_name:
                location = "torus side"
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
            value = f"{value_match.group(1)}" if value_match else "N/A"
            label = "Circularity" if tol_type.upper() == "ROUNDNESS" else tol_type.capitalize()

            # Enhanced datum letter and location detection
            datum_letter = ""
            location = ""
            
            # Method 1: Extract datum letter from tolerance name (e.g., "tolerance(A)")
            tol_name_lower = tol_name.lower()
            datum_match = re.search(r'\(([A-Z])\)', tol_name)
            if datum_match:
                datum_letter = datum_match.group(1)
                location = datum_results.get(datum_letter, "")
            
            # Method 2: Try to find datum letter from tolerance name with lowercase
            if not datum_letter:
                for d_letter in datum_results:
                    if f"({d_letter.lower()})" in tol_name_lower or f"({d_letter.upper()})" in tol_name:
                        datum_letter = d_letter
                        location = datum_results[d_letter]
                        break
            
            # Method 3: Check if tolerance name contains feature references
            if not location:
                # Look for Boss1, Plane1, Plane2 in tolerance name
                if "boss1" in tol_name_lower:
                    location = "cylindrical side"
                elif "plane1" in tol_name_lower:
                    location = "bottom face"
                elif "plane2" in tol_name_lower:
                    location = "top face"
                elif "plane" in tol_name_lower:
                    # Generic plane - try to determine from context
                    location = "planar surface"
            
            # Method 4: Infer location based on tolerance type if still not found
            if not location:
                if label.lower() in ["cylindricity", "circularity"]:
                    location = "cylindrical side"
                elif label.lower() in ["flatness"]:
                    location = "planar surface"
                elif label.lower() in ["straightness"]:
                    location = "surface"

            tol_results.append((label, value, datum_letter, location))

        # GD&T symbol mapping for output
        gdnt_symbols = {
            "Straightness": "─",
            "Flatness": "☐",
            "Circularity": "○",
            "Cylindricity": "⌀"
        }

        # Helper to clean and parse feature names - Enhanced for torus detection
        def clean_feature_name(feature_name):
            if not feature_name:
                return ""
            # Remove common STEP file artifacts and extract meaningful parts
            fname = str(feature_name).lower()
            # Remove datum references like "datum10@" or similar patterns
            fname = re.sub(r'datum\d+@', '', fname)
            # Extract geometric feature types
            if "torus" in fname:
                return "torus"
            elif "plane" in fname:
                return "plane"
            elif "boss" in fname:
                return "boss"
            elif "cylinder" in fname:
                return "cylinder"
            elif "cone" in fname:
                return "cone"
            else:
                # Keep original if no pattern matches
                return feature_name

        # Enhanced helper to map feature name to surface type (for Location column)
        def get_surface_type(feature_name):
            if not feature_name:
                return "surface"
                
            fname = str(feature_name).lower()
            
            # Check for torus patterns first
            if "torus" in fname:
                return "torus side"
            # Check for cleaned patterns
            cleaned_name = clean_feature_name(feature_name)
            if cleaned_name == "torus":
                return "torus side"
            
            # Enhanced plane detection for datums
            if "plane" in fname:
                return determine_plane_position(feature_name)
                
            if "cone" in fname or "conical" in fname:
                return "conical side of the part"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "cylindrical side"
            else:
                return str(feature_name)

        # Enhanced helper to map feature name and type to Surface (for Surface column)
        def get_likely_location(label, feature_name):
            if not feature_name:
                return "surface"
                
            fname = str(feature_name).lower()
            
            # Check for torus patterns first
            if "torus" in fname:
                return "torus side"
            # Check for cleaned patterns
            cleaned_name = clean_feature_name(feature_name)
            if cleaned_name == "torus":
                return "torus side"
            
            # Enhanced plane detection for datums
            if "plane" in fname:
                return determine_plane_position(feature_name)
                
            if "cone" in fname or "conical" in fname:
                return "conical side of the part"
            elif "boss1" in fname or "cylindrical" in fname or "side" in fname:
                return "curved side of the cylinder"
            elif "face" in fname:
                return "planar face"
            else:
                return str(feature_name)

        # Build table rows for results
        table_rows = []
        for label, value, datum, loc in tol_results:
            symbol = gdnt_symbols.get(label, "")
            type_with_symbol = f"{symbol} {label}" if symbol else label
            
            # Create location string with datum reference
            if datum:
                location_str = f"at datum {datum}"
            else:
                location_str = loc if loc else "surface"
            
            # Map location to surface description for consistency
            if loc == "cylindrical side":
                surface = "curved side of the cylinder"
            elif loc == "bottom face":
                surface = "bottom face"
            elif loc == "top face":
                surface = "top face"  
            elif loc == "planar surface":
                surface = "planar face"
            else:
                surface = loc if loc else "surface"
                
            table_rows.append(
                (type_with_symbol, value, datum, location_str, surface))
        
        # Enhanced datum processing with better plane detection and feature extraction
        for d_letter, feature_name in datum_letter_to_feature.items():
            location_str = datum_results.get(d_letter, "surface")
            surface = location_str
            
            # For datums, keep the original surface location in Location column
            table_rows.append(
                ("Datum", d_letter, d_letter, location_str, surface))
                
        return table_rows

    # Displays the results in a ttk.Treeview table
    def show_table(result_lines):
        # Destroy previous table if exists
        global output_table
        if output_table:
            output_table.destroy()
        columns = ("Type", "Value", "Datum", "Location", "Surface")
        output_table = ttk.Treeview(
            root, columns=columns, show="headings", height=20)
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
        rows = [output_table.item(row)['values']
                for row in output_table.get_children()]
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
            messagebox.showinfo(
                "Saved", f"Results saved as {file_format.upper()}!")
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
        filename_label.configure(
            bg=bg_color, fg="#bbbbbb" if dark_mode else "black")
        output_table.configure(bg="#1e1e1e" if dark_mode else "white",
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
    logo_frame = tk.Frame(top_frame, bg="#ffffff", borderwidth=1, relief="solid")
    logo_frame.grid(row=0, column=0, sticky="w", padx=10)
    try:
        logo_img = Image.open("./assets/unisza1.png").resize((100, 100))
        logo_photo = ImageTk.PhotoImage(logo_img)
        tk.Label(logo_frame, image=logo_photo, bg="#ffffff").pack()
    except:
        pass

    title_label = tk.Label(top_frame, text="GD&T Tolerance Value Extractor",
                           font=("Helvetica", 20, "bold"), bg="#eef3f7", fg="#003366")
    title_label.grid(row=0, column=1, padx=10, pady=10)

    right_logo_frame = tk.Frame(top_frame, bg="#ffffff", borderwidth=1, relief="solid")
    right_logo_frame.grid(row=0, column=2, sticky="e", padx=10)
    try:
        right_img = Image.open("./assets/frit.png").resize((100, 100))
        right_photo = ImageTk.PhotoImage(right_img)
        tk.Label(right_logo_frame, image=right_photo, bg="#ffffff").pack()
    except:
        pass

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
    tk.OptionMenu(mid_frame, format_var, ".txt", ".csv",
                  ".xlsx").grid(row=0, column=4, padx=5)

    filename_label = tk.Label(root, text="", font=(
        "Arial", 10), bg="#eef3f7", fg="black")
    filename_label.pack(pady=(5, 0))

    # Create empty table on startup
    global output_table
    columns = ("Type", "Value", "Datum", "Location", "Surface")
    output_table = ttk.Treeview(
        root, columns=columns, show="headings", height=20)
    for col in columns:
        output_table.heading(col, text=col)
        output_table.column(col, anchor="center", width=150)
    output_table.tag_configure('even', background='#f5f5f5')
    output_table.tag_configure('odd', background='#e9ecef')
    output_table.pack(expand=True, fill='both', padx=20, pady=10)

    root.mainloop()

create_interface()