#!/usr/bin/env python3
"""
Smart PDF Redactor - Enhanced Text Selection Edition with PII Detection
BY- CA. Vaibhav Balar

Ultra-high precision PDF redaction with intelligent text selection and PII detection
- Interactive text selection in PDF preview
- Smart pattern creation from selected text
- Context menu with pattern suggestions
- Copy-paste functionality 
- Advanced PII detection using NLP models
- Text de-identification capabilities
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
import re
import os
import io
import json
import webbrowser
import random
import sys
import glob
import shutil
import tempfile
from tkinter import simpledialog
import threading

# PII Detection imports
import os
import importlib.util

def check_import(module_name):
    """Check if a module can be imported"""
    return importlib.util.find_spec(module_name) is not None

# Check core dependencies with fallbacks
CORE_DEPENDENCIES = {
    'image_processing': ['PIL'],
    'pdf_processing': ['fitz', 'PyMuPDF'],
    'pii_detection': ['presidio_analyzer', 'presidio_anonymizer'],
    'nlp': ['spacy'],
    'advanced_nlp': ['transformers', 'flair']
}

DEPENDENCIES_STATUS = {group: False for group in CORE_DEPENDENCIES}

def check_dependency_group(group):
    """Check if any package in a dependency group is available"""
    return any(check_import(pkg) for pkg in CORE_DEPENDENCIES[group])

# Check each group
for group in CORE_DEPENDENCIES:
    DEPENDENCIES_STATUS[group] = check_dependency_group(group)

DEPENDENCIES_LOADED = all(DEPENDENCIES_STATUS.values())

if True:  # Always try to import what's available
    try:
        # Basic image processing
        if DEPENDENCIES_STATUS['image_processing']:
            from PIL import Image, ImageTk, ImageDraw, ImageFont
        
        # PDF processing
        if DEPENDENCIES_STATUS['pdf_processing']:
            try:
                import fitz  # Try PyMuPDF first
            except ImportError:
                import PyMuPDF as fitz  # Fallback name
        
        # PII detection if available
        if DEPENDENCIES_STATUS['pii_detection']:
            from presidio_analyzer import AnalyzerEngine
            from presidio_analyzer.nlp_engine import NlpEngineProvider
            from presidio_anonymizer import AnonymizerEngine
            from presidio_anonymizer.entities import OperatorConfig
        
        # NLP features if available
        if DEPENDENCIES_STATUS['nlp']:
            import spacy
            try:
                import en_core_web_lg
            except ImportError:
                pass  # Will download later if needed
        
        # Advanced NLP if available
        if DEPENDENCIES_STATUS['advanced_nlp']:
            try:
                from transformers import pipeline
            except ImportError:
                pass
                
            try:
                from flair.models import SequenceTagger
                from flair.data import Sentence
            except ImportError:
                pass
                
    except ImportError as e:
        print(f"Warning: Some imports failed - {str(e)}")
        # Continue with reduced functionality

def check_dependencies():
    """Check if all required dependencies are available and install if needed"""
    global DEPENDENCIES_LOADED, DEPENDENCIES_STATUS
    
    # Skip if everything is already loaded
    if DEPENDENCIES_LOADED:
        return DEPENDENCIES_LOADED
    
    # Check which features are missing
    missing_features = [group for group, status in DEPENDENCIES_STATUS.items() if not status]
    
    if missing_features and not hasattr(check_dependencies, '_already_asked'):
        check_dependencies._already_asked = True
        
        # Create a detailed message about missing features
        feature_descriptions = {
            'image_processing': 'Basic image handling',
            'pdf_processing': 'PDF reading and writing',
            'pii_detection': 'Personal information detection',
            'nlp': 'Text analysis',
            'advanced_nlp': 'Advanced language processing'
        }
        
        message = "Some features are not available:\n\n"
        for feature in missing_features:
            message += f"• {feature_descriptions[feature]}\n"
        message += "\nWould you like to install the required packages?"
        
        result = messagebox.askyesno("Missing Features", message)
        
        if result == 'yes':
            try:
                import subprocess
                import sys
                import os
                
                # Get pip path from current Python executable
                python_path = sys.executable
                pip_cmd = [python_path, "-m", "pip"]
                
                # Install packages in groups with error handling
                def install_single_package(package_name):
                    """Install a single package with detailed feedback"""
                    try:
                        messagebox.showinfo("Installing", f"Installing {package_name}...")
                        result = subprocess.run([*pip_cmd, "install", package_name], 
                                             stdout=subprocess.PIPE, 
                                             stderr=subprocess.PIPE,
                                             text=True)
                        
                        if result.returncode != 0:
                            error_msg = f"Failed to install {package_name}.\nError: {result.stderr}\n\nTry manually:\npip install {package_name}"
                            messagebox.showerror("Installation Error", error_msg)
                            return False
                        return True
                    except Exception as e:
                        error_msg = f"Error installing {package_name}:\n{str(e)}\n\nTry manually:\npip install {package_name}"
                        messagebox.showerror("Installation Error", error_msg)
                        return False

                def install_package_group(packages, group_name):
                    """Install packages one at a time with feedback"""
                    try:
                        # First upgrade pip
                        messagebox.showinfo("Upgrading pip", "Upgrading pip first...")
                        subprocess.run([*pip_cmd, "install", "--upgrade", "pip"],
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE)
                        
                        # Install each package individually
                        for package in packages:
                            if not install_single_package(package):
                                return False
                        return True
                    except Exception as e:
                        messagebox.showerror("Error",
                            f"Failed during installation of {group_name}:\n{str(e)}")
                        return False

                # Define packages by feature group
                install_groups = {
                    'image_processing': [
                        {"name": "Pillow", "desc": "Image processing library"}
                    ],
                    'pdf_processing': [
                        {"name": "PyMuPDF", "desc": "PDF processing library"}
                    ],
                    'pii_detection': [
                        {"name": "presidio-analyzer", "desc": "PII analyzer"},
                        {"name": "presidio-anonymizer", "desc": "PII anonymizer"}
                    ],
                    'nlp': [
                        {"name": "spacy", "desc": "NLP framework"}
                    ],
                    'advanced_nlp': [
                        {"name": "torch", "desc": "Deep learning framework"},
                        {"name": "transformers", "desc": "Transformer models"},
                        {"name": "flair==0.11.3", "desc": "NLP framework"}
                    ]
                }
                
                # Only try to install missing features
                success = True
                total_features = len(missing_features)
                
                if total_features > 0:
                    messagebox.showinfo("Installation Starting", 
                        "Will now install missing packages.\n" +
                        "This may take several minutes.")
                    
                    # Install each missing feature group
                    for feature in missing_features:
                        packages = install_groups[feature]
                        feature_msg = f"Installing packages for: {feature_descriptions[feature]}"
                        messagebox.showinfo("Installing Feature", feature_msg)
                        
                        # Try each package in the group
                        for package in packages:
                            package_name = package["name"]
                            description = package["desc"]
                            
                            # Show progress
                            progress_msg = f"Installing {package_name}\n({description})"
                            messagebox.showinfo("Package Installation", progress_msg)
                            
                            if not install_single_package(package_name):
                                retry = messagebox.askyesno("Installation Failed",
                                    f"Failed to install {package_name}.\nWould you like to skip this package and continue?")
                                if not retry:
                                    success = False
                                    break
                            
                # Handle spaCy model separately
                if success:
                    messagebox.showinfo("Installing spaCy Model",
                        "Now installing English language model...\nThis may take a few minutes.")
                    try:
                        result = subprocess.run([python_path, "-m", "spacy", "download", "en_core_web_lg"],
                                             stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                        if result.returncode != 0:
                            error_msg = f"Failed to install spaCy model.\nError: {result.stderr}"
                            messagebox.showerror("Model Installation Error", error_msg)
                            success = False
                    except Exception as e:
                        error_msg = f"Error installing spaCy model:\n{str(e)}"
                        messagebox.showerror("Model Installation Error", error_msg)
                        success = False

                if not success:
                    # Try alternative installation method
                    try:
                        subprocess.check_call([*pip_cmd, "install", "--user", 
                            "torch", "spacy", "transformers", "flair==0.11.3",
                            "presidio-analyzer", "presidio-anonymizer",
                            "--no-deps"])
                        return True
                    except subprocess.CalledProcessError as e:
                        messagebox.showerror("Error",
                            "Failed to install dependencies with alternative method.\n" +
                            "Please try installing manually:\n" +
                            "1. Open command prompt as administrator\n" +
                            "2. Run: pip install torch spacy transformers flair==0.11.3 presidio-analyzer presidio-anonymizer")
                        return False
                
                # Install spaCy model with better error handling
                try:
                    messagebox.showinfo("Installing", "Installing spaCy language model...\nThis may take a few minutes.")
                    result = subprocess.run([python_path, "-m", "spacy", "download", "en_core_web_lg"],
                                         stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                         text=True)
                    if result.returncode != 0:
                        messagebox.showerror("Error",
                            "Failed to install spaCy model. Error message:\n" + result.stderr)
                        return False
                except Exception as e:
                    messagebox.showerror("Error",
                        f"Failed to install spaCy model:\n{str(e)}\n\nTry running:\npip install spacy\npython -m spacy download en_core_web_lg")
                    return False
                    
                messagebox.showinfo("Success", 
                    "Dependencies installed successfully.\n" +
                    "Please restart the application.")
                sys.exit(0)
                
            except Exception as e:
                messagebox.showerror("Error",
                    f"Failed to install dependencies:\n{str(e)}")
                return False
        return False
    return True

# Suppress console output for EXE deployment
if getattr(sys, 'frozen', False):
    import sys
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.devnull, 'w')

# Enhanced predefined patterns with ultra-precise regex
PREDEFINED_PATTERNS = {
    "PAN Card Number": r"\b[A-Z]{5}[0-9]{4}[A-Z]\b",
    "GST Number": r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b",
    "Email Address": r"\b[a-zA-Z0-9]([a-zA-Z0-9._-]*[a-zA-Z0-9])?@[a-zA-Z0-9]([a-zA-Z0-9.-]*[a-zA-Z0-9])?\.[a-zA-Z]{2,}\b",
    "Phone Number (10 digits)": r"\b\d{10}\b",
    "Amounts ending with .00": r"\d+\.00\b",
    "Aadhaar Number": r"\b\d{4}\s\d{4}\s\d{4}\b",
    "Passport Number": r"\b[A-PR-WY][1-9]\d{6}\b",
    "Driving License": r"\b[A-Z]{2}\d{2}\s\d{11}\b",
    "Voter ID": r"\b[A-Z]{3}[0-9]{7}\b",
    "CIN Number": r"\b[LU][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}\b",
    "Bank Account Number": r"\b\d{9,18}\b",
    "IFSC Code": r"\b[A-Z]{4}0[A-Z0-9]{6}\b",
    "UPI ID": r"\b[a-zA-Z0-9.\-_]{2,}@[a-zA-Z]{3,}\b",
    "Credit Card Number": r"\b(?:\d[ -]*?){13,16}\b",
    "CVV (3 digits)": r"\b\d{3}\b",
    "TAN Number": r"\b[A-Z]{4}\d{5}[A-Z]\b",
    "IP Address (IPv4)": r"\b(?:\d{1,3}\.){3}\d{1,3}\b",
    "URL": r"https?://(?:www\.)?[a-zA-Z0-9]([a-zA-Z0-9.-]*[a-zA-Z0-9])?(?:\.[a-zA-Z]{2,})+(?:/[^\s]*)?|www\.[a-zA-Z0-9]([a-zA-Z0-9.-]*[a-zA-Z0-9])?\.[a-zA-Z]{2,}(?:/[^\s]*)?",
    "INR Amount (₹ Format)": r"₹\s?\d{1,3}(,\d{3})*(\.\d{2})?",
    "PNR Number": r"\b\d{10}\b",
    "Indian Mobile Number": r"\b[6-9]\d{9}\b",
    "Date (DD/MM/YYYY)": r"\b\d{2}/\d{2}/\d{4}\b",
    "Date (YYYY-MM-DD)": r"\b\d{4}-\d{2}-\d{2}\b",
    "Pincode": r"\b\d{6}\b",
    "Vehicle Number": r"\b[A-Z]{2}\d{2}[A-Z]{2}\d{4}\b",
    "Invoice Number": r"\b(?:INV|INVOICE)[-\s]*\d+\b",
    "Order Number": r"\b(?:ORD|ORDER)[-\s]*\d+\b"
}

# NLP Model configurations
NLP_MODELS = {
    "spaCy/en_core_web_lg": {
        "description": "General-purpose NER model",
        "type": "spacy",
        "model": "en_core_web_lg",
        "entities": ["PERSON", "ORG", "GPE", "LOC", "PRODUCT", "EVENT", "DATE", "TIME", "MONEY"]
    },
    "flair/ner-english-large": {
        "description": "High-accuracy NER model",
        "type": "flair",
        "model": "flair/ner-english-large",
        "entities": ["PER", "ORG", "LOC", "MISC"]
    },
    "HuggingFace/deid_roberta": {
        "description": "Medical data de-identification",
        "type": "transformers",
        "model": "obi/deid_roberta_i2b2",
        "entities": ["PATIENT", "DOCTOR", "HOSPITAL", "MEDICALRECORD", "DATE"]
    }
}

# PII Entity Types with descriptions
PII_ENTITY_TYPES = {
    "PERSON": {"description": "Names of persons"},
    "LOCATION": {"description": "Addresses, cities, countries"},
    "ORGANIZATION": {"description": "Company and organization names"},
    "EMAIL": {"description": "Email addresses"},
    "PHONE_NUMBER": {"description": "Phone and fax numbers"},
    "CREDIT_CARD": {"description": "Credit card numbers"},
    "CRYPTO": {"description": "Cryptocurrency addresses"},
    "IP_ADDRESS": {"description": "IP addresses"},
    "BANK_ACCOUNT": {"description": "Bank account numbers"},
    "ID_NUMBER": {"description": "Various ID numbers"},
    "DATE_TIME": {"description": "Dates and times"},
    "US_SSN": {"description": "US Social Security numbers"},
    "US_PASSPORT": {"description": "US Passport numbers"},
    "MEDICAL_RECORD": {"description": "Medical record numbers"}
}

# Masking styles
MASKING_STYLES = {
    "type_label": {"description": "Replace with <TYPE>"},
    "full_mask": {"description": "Replace with ****"},
    "partial_mask": {"description": "Show first/last characters"},
    "redact": {"description": "Replace with black box"}
}

# Advanced match types
ADVANCED_MATCH_TYPES = {
    "Exact word match": {
        "description": "Match exact word boundaries",
        "inputs": ["Word"],
        "generator": lambda word: rf"\b{re.escape(word)}\b"
    },
    "Text after word": {
        "description": "Text that comes immediately after a specific word",
        "inputs": ["Word"],
        "generator": lambda word: rf"(?<={re.escape(word)}\s+)\S+"
    },
    "Text before word": {
        "description": "Text that comes immediately before a specific word", 
        "inputs": ["Word"],
        "generator": lambda word: rf"\S+(?=\s+{re.escape(word)})"
    },
    "Line after word": {
        "description": "Entire line that comes after a word",
        "inputs": ["Word"],
        "generator": lambda word: rf"(?<={re.escape(word)}).*"
    },
    "Word starts with": {
        "description": "Words starting with specific text",
        "inputs": ["Prefix"],
        "generator": lambda prefix: rf"\b{re.escape(prefix)}\w*"
    },
    "Word ends with": {
        "description": "Words ending with specific text",
        "inputs": ["Suffix"], 
        "generator": lambda suffix: rf"\b\w*{re.escape(suffix)}\b"
    },
    "Word contains": {
        "description": "Words containing specific text anywhere",
        "inputs": ["Text"],
        "generator": lambda text: rf"\b\w*{re.escape(text)}\w*\b"
    },
    "Between two words": {
        "description": "Text between two specific words",
        "inputs": ["Start Word", "End Word"],
        "generator": lambda start, end: rf"(?<={re.escape(start)}).*?(?={re.escape(end)})"
    },
    "Entire line containing": {
        "description": "Complete line containing specific text",
        "inputs": ["Text"],
        "generator": lambda text: rf"^.*{re.escape(text)}.*$"
    },
    "Text in parentheses": {
        "description": "Any text enclosed in parentheses",
        "inputs": [],
        "generator": lambda: r"$$[^)]+$$"
    },
    "Text in quotes": {
        "description": "Any text enclosed in double quotes",
        "inputs": [],
        "generator": lambda: r'"[^"]*"'
    },
    "Numbers after text": {
        "description": "Numbers that follow specific text",
        "inputs": ["Text"],
        "generator": lambda text: rf"(?<={re.escape(text)}\s*)\d+"
    },
    "Text after colon": {
        "description": "Text that comes after a colon",
        "inputs": ["Label"],
        "generator": lambda label: rf"(?<={re.escape(label)}:\s*)[^\n\r]+"
    },
    "Custom Regex": {
        "description": "Enter your own regular expression pattern",
        "inputs": ["Regex Pattern"],
        "generator": lambda pattern: pattern
    },
    "Single character": {
        "description": "Match any occurrence of a specific character or letter",
        "inputs": ["Character"],
        "generator": lambda char: re.escape(char)
    },
    "Multiple characters": {
        "description": "Match any of the specified characters",
        "inputs": ["Characters (e.g., ABC123)"],
        "generator": lambda chars: f"[{re.escape(chars)}]"
    },
    "All digits": {
        "description": "Match any single digit (0-9)",
        "inputs": [],
        "generator": lambda: r"\d"
    },
    "All letters": {
        "description": "Match any single letter (A-Z, a-z)",
        "inputs": [],
        "generator": lambda: r"[A-Za-z]"
    },
    "Uppercase letters": {
        "description": "Match any uppercase letter (A-Z)",
        "inputs": [],
        "generator": lambda: r"[A-Z]"
    },
    "Lowercase letters": {
        "description": "Match any lowercase letter (a-z)",
        "inputs": [],
        "generator": lambda: r"[a-z]"
    },
}

class AutocompleteCombobox(ttk.Combobox):
    """Optimized Combobox with autocomplete functionality"""
    def __init__(self, parent, values=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.original_values = values or []
        self['values'] = self.original_values
        self.var = self['textvariable']
        if not self.var:
            self.var = tk.StringVar()
            self['textvariable'] = self.var
        
        # Enhanced event binding for autocomplete
        self.bind('<KeyRelease>', self.on_keyrelease)
        self.bind('<Button-1>', self.on_click)
        self.bind('<FocusIn>', self.on_focus_in)
        
    def on_keyrelease(self, event):
        """Enhanced keyrelease handler with working autocomplete"""
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape']:
            return
            
        typed = self.get().lower()
        if typed == '':
            self['values'] = self.original_values
            return
            
        # Enhanced matching - starts with and contains
        starts_with = [item for item in self.original_values if item.lower().startswith(typed)]
        contains = [item for item in self.original_values if typed in item.lower() and item not in starts_with]
        
        # Combine results - starts_with first, then contains
        matches = starts_with + contains
        self['values'] = matches[:15]  # Show more options
        
        # Auto-complete first match
        if matches and len(typed) > 0:
            first_match = matches[0]
            if first_match.lower().startswith(typed):
                # Set the text and select the auto-completed part
                self.delete(0, tk.END)
                self.insert(0, first_match)
                self.select_range(len(typed), tk.END)
                self.icursor(len(typed))
        
    def on_click(self, event):
        """Show all values when clicked"""
        self['values'] = self.original_values
        
    def on_focus_in(self, event):
        """Show all values when focused"""
        self['values'] = self.original_values
        
    def update_values(self, new_values):
        """Update the list of values"""
        self.original_values = new_values
        self['values'] = new_values

class CollapsibleFrame:
    """Lightweight collapsible frame"""
    def __init__(self, parent, title, icon="📁", bg="#ffffff"):
        self.parent = parent
        self.title = title
        self.icon = icon
        self.bg = bg
        self.is_expanded = True
        
        # Main container
        self.container = tk.Frame(parent, bg=bg)
        self.container.pack(fill=tk.X, padx=4, pady=4)
        
        # Header frame
        self.header_frame = tk.Frame(self.container, bg="#f0f0f0", relief=tk.RAISED, bd=1)
        self.header_frame.pack(fill=tk.X)
        
        # Toggle button
        self.toggle_btn = tk.Button(
            self.header_frame, 
            text=f"▼ {icon} {title}",
            command=self.toggle,
            bg="#f0f0f0", 
            fg="#1a237e",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT,
            anchor="w",
            padx=8,
            pady=4,
            cursor="hand2"
        )
        self.toggle_btn.pack(fill=tk.X)
        
        # Content frame
        self.content_frame = tk.Frame(self.container, bg=bg, relief=tk.FLAT, bd=1)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
    def toggle(self):
        """Toggle visibility"""
        if self.is_expanded:
            self.content_frame.pack_forget()
            self.toggle_btn.config(text=f"▶ {self.icon} {self.title}")
            self.is_expanded = False
        else:
            self.content_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
            self.toggle_btn.config(text=f"▼ {self.icon} {self.title}")
            self.is_expanded = True
    
    def get_content_frame(self):
        """Get the content frame for adding widgets"""
        return self.content_frame

class OptimizedScrollableFrame:
    """Lightweight scrollable frame with better performance and smooth scrolling"""
    def __init__(self, parent, **kwargs):
        self.outer_frame = tk.Frame(parent, **kwargs)
        
        # Create canvas
        self.canvas = tk.Canvas(
            self.outer_frame, 
            highlightthickness=0, 
            bg=kwargs.get('bg', '#ffffff'),
            bd=0,
            relief='flat'
        )
        
        # Create scrollbar
        self.v_scrollbar = ttk.Scrollbar(
            self.outer_frame, 
            orient="vertical", 
            command=self.canvas.yview
        )
        
        # Configure canvas scrolling
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)
        
        # Create inner frame
        self.inner_frame = tk.Frame(self.canvas, bg=kwargs.get('bg', '#ffffff'))
        self.canvas_window = self.canvas.create_window(
            (0, 0), 
            window=self.inner_frame, 
            anchor="nw"
        )
        
        # Pack elements
        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scrollbar.pack(side="right", fill="y")
        
        # Bind events
        self.inner_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # IMPROVED MOUSE WHEEL BINDING - bind to multiple widgets
        self.bind_mousewheel_to_widget(self.canvas)
        self.bind_mousewheel_to_widget(self.inner_frame)
        self.bind_mousewheel_to_widget(self.outer_frame)
        
    def bind_mousewheel_to_widget(self, widget):
        """Bind mouse wheel to widget for smooth scrolling"""
        widget.bind("<MouseWheel>", self.on_mousewheel)
        widget.bind("<Button-4>", self.on_mousewheel)  # Linux
        widget.bind("<Button-5>", self.on_mousewheel)  # Linux
        
        # Bind to all child widgets recursively
        def bind_to_children(parent):
            for child in parent.winfo_children():
                try:
                    child.bind("<MouseWheel>", self.on_mousewheel)
                    child.bind("<Button-4>", self.on_mousewheel)
                    child.bind("<Button-5>", self.on_mousewheel)
                    bind_to_children(child)
                except:
                    pass
        
        widget.after(100, lambda: bind_to_children(widget))
        
    def on_mousewheel(self, event):
        """Improved mouse wheel scrolling with better sensitivity"""
        try:
            # Handle different platforms
            if event.delta:
                delta = -1 * (event.delta / 120)
            elif event.num == 4:
                delta = -1
            elif event.num == 5:
                delta = 1
            else:
                delta = 0
            
            # Smooth scrolling
            self.canvas.yview_scroll(int(delta), "units")
        except:
            pass
        
    def on_frame_configure(self, event):
        """Update scroll region when frame size changes"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def on_canvas_configure(self, event):
        """Update inner frame width when canvas size changes"""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
        
    def pack(self, **kwargs):
        self.outer_frame.pack(**kwargs)
        
    def update_scroll_region(self):
        """Manually update scroll region"""
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        # Re-bind mouse wheel after updates
        self.bind_mousewheel_to_widget(self.inner_frame)

class TextSelectionHandler:
    """Enhanced text selection like Adobe Reader with drag and shift selection"""
    
    def __init__(self, parent_app):
        self.parent_app = parent_app
        self.selected_text = ""
        self.context_menu = None
        self.selection_start = None
        self.selection_end = None
        self.is_selecting = False
        self.selection_rect_id = None
        
    def setup_selection_bindings(self, canvas):
        """Setup enhanced selection bindings"""
        canvas.bind("<Double-Button-1>", self.on_double_click)
        canvas.bind("<Button-1>", self.on_single_click)
        canvas.bind("<B1-Motion>", self.on_drag)
        canvas.bind("<ButtonRelease-1>", self.on_release)
        canvas.bind("<Shift-Button-1>", self.on_shift_click)
        
    def on_single_click(self, event):
        """Handle single click to start selection"""
        canvas = event.widget
        canvas_x = canvas.canvasx(event.x)
        canvas_y = canvas.canvasy(event.y)
        
        # Convert to PDF coordinates
        pdf_x = canvas_x / self.parent_app.zoom
        pdf_y = canvas_y / self.parent_app.zoom
        
        self.selection_start = (pdf_x, pdf_y)
        self.selection_end = None
        self.is_selecting = True
        self.selected_text = ""
        
        # Clear previous selection rectangle
        if self.selection_rect_id:
            canvas.delete(self.selection_rect_id)
            self.selection_rect_id = None
    
    def on_drag(self, event):
        """Handle drag to extend selection"""
        if not self.is_selecting or not self.selection_start:
            return
            
        canvas = event.widget
        canvas_x = canvas.canvasx(event.x)
        canvas_y = canvas.canvasy(event.y)
        
        # Convert to PDF coordinates
        pdf_x = canvas_x / self.parent_app.zoom
        pdf_y = canvas_y / self.parent_app.zoom
        
        self.selection_end = (pdf_x, pdf_y)
        
        # Draw selection rectangle
        self.draw_selection_rectangle(canvas)
    
    def on_release(self, event):
        """Handle mouse release to finalize selection"""
        if not self.is_selecting:
            return
            
        self.is_selecting = False
        
        if self.selection_start and self.selection_end:
            # Get text in selection area
            self.get_text_in_selection_area()
            
            if self.selected_text.strip():
                self.show_simple_context_menu(event)
    
    def on_shift_click(self, event):
        """Handle shift+click to extend selection"""
        if not self.selection_start:
            self.on_single_click(event)
            return
            
        canvas = event.widget
        canvas_x = canvas.canvasx(event.x)
        canvas_y = canvas.canvasy(event.y)
        
        # Convert to PDF coordinates
        pdf_x = canvas_x / self.parent_app.zoom
        pdf_y = canvas_y / self.parent_app.zoom
        
        self.selection_end = (pdf_x, pdf_y)
        self.get_text_in_selection_area()
        
        if self.selected_text.strip():
            self.show_simple_context_menu(event)
    
    def draw_selection_rectangle(self, canvas):
        """Draw selection rectangle on canvas"""
        if not self.selection_start or not self.selection_end:
            return
            
        # Clear previous rectangle
        if self.selection_rect_id:
            canvas.delete(self.selection_rect_id)
        
        # Convert PDF coordinates to canvas coordinates
        start_x = self.selection_start[0] * self.parent_app.zoom
        start_y = self.selection_start[1] * self.parent_app.zoom
        end_x = self.selection_end[0] * self.parent_app.zoom
        end_y = self.selection_end[1] * self.parent_app.zoom
        
        # Ensure proper rectangle coordinates
        x1, x2 = min(start_x, end_x), max(start_x, end_x)
        y1, y2 = min(start_y, end_y), max(start_y, end_y)
        
        # Draw selection rectangle
        self.selection_rect_id = canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="#2196f3", width=1, fill="#2196f3",
            stipple="gray50", tags="selection"
        )
    
    def get_text_in_selection_area(self):
        """Get text within the selection area"""
        try:
            if not self.parent_app.pdf_path or not self.selection_start or not self.selection_end:
                return
                
            doc = fitz.open(self.parent_app.pdf_path)
            page = doc[self.parent_app.current_page]
            
            # Create selection rectangle
            x1, x2 = min(self.selection_start[0], self.selection_end[0]), max(self.selection_start[0], self.selection_end[0])
            y1, y2 = min(self.selection_start[1], self.selection_end[1]), max(self.selection_start[1], self.selection_end[1])
            
            selection_rect = fitz.Rect(x1, y1, x2, y2)
            
            # Get text within selection area
            words = page.get_text("words")
            selected_words = []
            
            for word_data in words:
                word_x0, word_y0, word_x1, word_y1, word_text, block_no, line_no, word_no = word_data
                word_rect = fitz.Rect(word_x0, word_y0, word_x1, word_y1)
                
                # Check if word intersects with selection
                if selection_rect.intersects(word_rect):
                    # For partial word selection, calculate character-level selection
                    if selection_rect.contains(word_rect):
                        # Entire word is selected
                        selected_words.append(word_text)
                    else:
                        # Partial word selection - estimate characters
                        overlap = selection_rect & word_rect
                        if overlap.width > 0 and overlap.height > 0:
                            # Calculate which part of the word is selected
                            word_width = word_rect.width
                            char_width = word_width / len(word_text) if len(word_text) > 0 else word_width
                            
                            start_char = max(0, int((overlap.x0 - word_rect.x0) / char_width))
                            end_char = min(len(word_text), int((overlap.x1 - word_rect.x0) / char_width) + 1)
                            
                            if start_char < end_char:
                                selected_words.append(word_text[start_char:end_char])
            
            self.selected_text = ' '.join(selected_words).strip()
            doc.close()
                        
        except Exception as e:
            print(f"Error getting selection text: {e}")
            self.selected_text = ""
        
    def on_double_click(self, event):
        """Handle double click for word selection - Adobe Reader style"""
        canvas = event.widget
        canvas_x = canvas.canvasx(event.x)
        canvas_y = canvas.canvasy(event.y)
        
        # Convert to PDF coordinates
        pdf_x = canvas_x / self.parent_app.zoom
        pdf_y = canvas_y / self.parent_app.zoom
        
        # Select word at click position
        self.select_word_at_position(pdf_x, pdf_y)
        
        if self.selected_text.strip():
            self.show_simple_context_menu(event)
    
    def select_word_at_position(self, pdf_x, pdf_y):
        """Select word at the given PDF coordinates - IMPROVED"""
        try:
            if not self.parent_app.pdf_path:
                return
                
            doc = fitz.open(self.parent_app.pdf_path)
            page = doc[self.parent_app.current_page]
            
            # Use PyMuPDF's built-in word detection
            words = page.get_text("words")  # Get word-level data
            
            # Find the word at the clicked position
            for word_data in words:
                x0, y0, x1, y1, word_text, block_no, line_no, word_no = word_data
                word_rect = fitz.Rect(x0, y0, x1, y1)
                
                # Check if click is within this word
                if word_rect.contains(fitz.Point(pdf_x, pdf_y)):
                    self.selected_text = word_text.strip()
                    doc.close()
                    return
            
            doc.close()
            self.selected_text = ""
                        
        except Exception as e:
            print(f"Error selecting word: {e}")
            self.selected_text = ""
    
    def show_simple_context_menu(self, event):
        """Show simple context menu - Adobe Reader style"""
        if not self.selected_text.strip():
            return
        
        # Destroy existing context menu
        if self.context_menu:
            self.context_menu.destroy()
    
        # Create simple context menu
        self.context_menu = tk.Menu(self.parent_app.window, tearoff=0)
    
        # Add copy option
        self.context_menu.add_command(
            label=f"📋 Copy '{self.selected_text[:20]}{'...' if len(self.selected_text) > 20 else ''}'",
            command=self.copy_selected_text
        )
    
        self.context_menu.add_separator()
    
        # Add header
        self.context_menu.add_command(
            label=f"Create Pattern with '{self.selected_text[:15]}{'...' if len(self.selected_text) > 15 else ''}':",
            state="disabled"
        )
    
        # Add custom pattern options from ADVANCED_MATCH_TYPES
        for pattern_type in ADVANCED_MATCH_TYPES.keys():
            self.context_menu.add_command(
                label=f"➕ {pattern_type}",
                command=lambda pt=pattern_type: self.create_pattern_from_selection(pt)
        )
    
        # Show context menu
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        except:
            pass
        finally:
            self.context_menu.grab_release()
    
    def copy_selected_text(self):
        """Copy selected text to clipboard"""
        if self.selected_text.strip():
            self.parent_app.window.clipboard_clear()
            self.parent_app.window.clipboard_append(self.selected_text)
            self.parent_app.status_text.set(f"📋 Copied: '{self.selected_text[:30]}{'...' if len(self.selected_text) > 30 else ''}'")
    
    def create_pattern_from_selection(self, pattern_type):
        """Create a pattern from selected text - DIRECT CREATION"""
        if not self.selected_text.strip():
            return
        
        try:
            # Get pattern info
            if pattern_type not in ADVANCED_MATCH_TYPES:
                return
            
            pattern_info = ADVANCED_MATCH_TYPES[pattern_type]
        
            # Generate regex and label directly
            if len(pattern_info["inputs"]) == 0:
                regex = pattern_info["generator"]()
                label = pattern_type
            elif len(pattern_info["inputs"]) == 1:
                regex = pattern_info["generator"](self.selected_text)
                label = f"{pattern_type}: {self.selected_text}"
            elif len(pattern_info["inputs"]) == 2:
                # For patterns that need two inputs, use selected text as first input
                # and prompt for second input
                second_input = tk.simpledialog.askstring(
                    "Additional Input", 
                    f"Enter {pattern_info['inputs'][1]} for '{self.selected_text}':",
                    parent=self.parent_app.window
                )
                if not second_input:
                    return
                regex = pattern_info["generator"](self.selected_text, second_input)
                label = f"{pattern_type}: {self.selected_text} + {second_input}"
            else:
                regex = pattern_info["generator"](self.selected_text)
                label = f"{pattern_type}: {self.selected_text}"
        
            # Check for duplicates
            for existing_pattern in self.parent_app.patterns:
                if existing_pattern["label"] == label:
                    messagebox.showwarning("Warning", f"Pattern '{label}' already exists.")
                    return
        
            color = self.parent_app.generate_random_color()
        
            new_pattern = {
                "label": label,
                "regex": regex,
                "color": color,
                "type": "selection"
            }
        
            # Add to patterns
            self.parent_app.patterns.append(new_pattern)
            self.parent_app.add_pattern_display(color, label, len(self.parent_app.patterns) - 1, regex)
            self.parent_app.update_pattern_count()
            self.parent_app.status_text.set(f"➕ Added pattern from selection: {label}")
        
            # AUTO-PREVIEW the new pattern
            if self.parent_app.pdf_path:
                self.parent_app.ultra_precise_preview()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create pattern: {str(e)}")

class PIIDetector:
    """Handles PII detection using various NLP models"""
    
    def __init__(self):
        self.models = {}
        self.current_model = None
        self.analyzer = None
        self.anonymizer = None
        self.threshold = 0.35
        self.is_initialized = False
        
        # Disable PII detection if dependencies aren't available
        if not DEPENDENCIES_LOADED:
            self.is_available = lambda: False
            return
        
    def is_available(self):
        """Check if PII detection is available"""
        return DEPENDENCIES_LOADED
        
    def ensure_initialized(self):
        """Ensure the detector is initialized"""
        if not self.is_initialized and DEPENDENCIES_LOADED:
            try:
                # Initialize Presidio components
                self.analyzer = AnalyzerEngine()
                self.anonymizer = AnonymizerEngine()
                self.is_initialized = True
                return True
            except Exception as e:
                print(f"Error initializing PII detector: {e}")
                return False
        return self.is_initialized
        
    def initialize_models(self, model_name):
        """Initialize selected NLP model"""
        try:
            print(f"\nInitializing model: {model_name}")
            self.current_model_name = model_name
            
            if model_name not in self.models:
                if model_name.startswith("spaCy"):
                    print("Loading spaCy model...")
                    try:
                        self.models[model_name] = spacy.load("en_core_web_lg")
                    except OSError:
                        print("Downloading spaCy model...")
                        spacy.cli.download("en_core_web_lg")
                        self.models[model_name] = spacy.load("en_core_web_lg")
                        
                elif model_name.startswith("flair"):
                    try:
                        print("Loading Flair model...")
                        print("Creating Flair SequenceTagger...")
                        model = SequenceTagger.load("flair/ner-english-large")
                        print("SequenceTagger created successfully")
                        self.models[model_name] = model
                        print("Model stored in models dictionary")
                    except Exception as e:
                        print(f"Error loading Flair model: {str(e)}")
                        raise
                    
                elif model_name.startswith("HuggingFace"):
                    try:
                        print("Loading Transformer model...")
                        model = pipeline("ner", 
                                      model="obi/deid_roberta_i2b2",
                                      aggregation_strategy="simple")
                        print("Pipeline created successfully")
                        self.models[model_name] = model
                    except Exception as e:
                        print(f"Error loading Transformer model: {str(e)}")
                        raise
            
            print("Setting current model...")
            self.current_model = self.models[model_name]
            print(f"Model loaded successfully: {type(self.current_model).__name__}")
            
            # Initialize Presidio components
            print("Initializing Presidio components...")
            self.analyzer = AnalyzerEngine()
            self.anonymizer = AnonymizerEngine()
            print("Presidio components initialized")
            
            return True
            
        except Exception as e:
            error_msg = f"Failed to initialize model: {str(e)}"
            print(f"Error: {error_msg}")
            messagebox.showerror("Model Initialization Error", error_msg)
            return False
            
            # Initialize Presidio
            self.analyzer = AnalyzerEngine()
            self.anonymizer = AnonymizerEngine()
            
            return True
        except Exception as e:
            print(f"Error initializing model {model_name}: {e}")
            return False
            
    def detect_entities(self, text, selected_types=None):
        """Detect PII entities in text using current model"""
        try:
            print("\n=== Starting Entity Detection ===")
            
            if not self.current_model:
                print("ERROR: No model loaded")
                messagebox.showerror("Error", "No model is loaded. Please select a model first.")
                return []
                
            if not text.strip():
                print("ERROR: Empty text")
                messagebox.showwarning("Warning", "Please enter some text to analyze.")
                return []
                
            entities = []
            
            # Debug info
            print(f"\nInput Parameters:")
            print(f"- Text: '{text}'")
            print(f"- Selected types: {selected_types}")
            print(f"- Model: {type(self.current_model).__name__}")
            print(f"- Model name: {self.current_model_name}")
            
            # Get model type
            print("\nGetting model info...")
            model_info = next((info for name, info in NLP_MODELS.items() 
                             if name == self.current_model_name), None)
            
            if not model_info:
                error_msg = f"Model info not found for {self.current_model_name}"
                print(f"ERROR: {error_msg}")
                messagebox.showerror("Error", error_msg)
                return []
                
            print(f"Model type: {model_info['type']}")
            print(f"Model description: {model_info['description']}")
            print(f"Supported entities: {model_info['entities']}")
                
            model_type = model_info["type"]
            print(f"Using model type: {model_type}")
            
            # Use appropriate detection method based on model type
            if model_type == "spacy":
                doc = self.current_model(text)
                entities = [{"entity": ent.label_,
                           "start": ent.start_char,
                           "end": ent.end_char,
                           "text": ent.text,
                           "score": 0.8}  # spaCy doesn't provide confidence scores
                          for ent in doc.ents]
                print(f"spaCy found {len(doc.ents)} entities")
                          
            elif model_type == "flair":
                try:
                    sentence = Sentence(text)
                    self.current_model.predict(sentence)
                    print(f"Flair prediction completed")
                    entities = []
                    for ent in sentence.get_spans('ner'):
                        entity = {
                            "entity": ent.tag,
                            "start": ent.start_pos,
                            "end": ent.end_pos,
                            "text": ent.text,
                            "score": ent.score
                        }
                        print(f"Found entity: {entity}")
                        entities.append(entity)
                except Exception as e:
                    print(f"Flair error: {str(e)}")
                    return []
                          
            elif model_type == "transformers":
                try:
                    results = self.current_model(text)
                    print(f"Transformer returned {len(results)} results")
                    entities = []
                    for ent in results:
                        entity = {
                            "entity": ent["entity_group"],
                            "start": ent["start"],
                            "end": ent["end"],
                            "text": ent["word"],
                            "score": ent["score"]
                        }
                        print(f"Found entity: {entity}")
                        entities.append(entity)
                except Exception as e:
                    print(f"Transformer error: {str(e)}")
                    return []
            
            print(f"Total entities before filtering: {len(entities)}")
            
            # Filter by selected types and threshold
            if selected_types:
                filtered_entities = []
                for ent in entities:
                    if ent["entity"] in selected_types and ent["score"] >= self.threshold:
                        filtered_entities.append(ent)
                        print(f"Keeping entity: {ent}")
                    else:
                        print(f"Filtering out entity: {ent}")
                entities = filtered_entities
            
            print(f"Final entities after filtering: {len(entities)}")
            return entities
            
        except Exception as e:
            print(f"Error detecting entities: {e}")
            return []
            
    def mask_entities(self, text, entities, style="type_label"):
        """Mask detected entities using specified style"""
        try:
            if not entities:
                return text
                
            print(f"\nMasking entities with style: {style}")
            print(f"Input text: '{text}'")
            print(f"Entities to mask: {entities}")
            
            # Validate inputs
            if not isinstance(text, str):
                raise ValueError(f"Text must be a string, got {type(text)}")
                
            # Create a list of all positions that need masking
            masks = []
            for ent in entities:
                try:
                    start = int(ent["start"])
                    end = int(ent["end"])
                    entity_type = str(ent["entity"])
                    
                    # Validate indices
                    if not (0 <= start < len(text) and 0 < end <= len(text) and start < end):
                        print(f"Invalid indices: start={start}, end={end}, text length={len(text)}")
                        continue
                        
                    entity_text = text[start:end]
                    print(f"\nProcessing entity: '{entity_text}' ({entity_type})")
                    print(f"Position: {start}-{end}")
                    
                    if style == "type_label":
                        mask_text = f"<{entity_type}>"
                    elif style == "full_mask":
                        mask_text = "*" * len(entity_text)
                    elif style == "partial_mask":
                        if len(entity_text) > 4:
                            mask_text = entity_text[0] + "*" * (len(entity_text) - 2) + entity_text[-1]
                        else:
                            mask_text = "*" * len(entity_text)
                    else:
                        mask_text = "****"
                        
                    print(f"Mask text: '{mask_text}'")
                    masks.append((start, end, mask_text))
                    
                except (KeyError, ValueError, IndexError) as e:
                    print(f"Invalid entity format: {e}")
                    continue
                except Exception as e:
                    print(f"Error processing entity: {e}")
                    continue
            
            if not masks:
                return text
                
            # Sort masks in reverse order to handle overlapping entities
            masks.sort(key=lambda x: x[0], reverse=True)
            print(f"\nApplying {len(masks)} masks")
            
            # Apply masks
            masked_text = text
            for start, end, mask_text in masks:
                try:
                    masked_text = masked_text[:start] + mask_text + masked_text[end:]
                except Exception as e:
                    print(f"Error applying mask: {e}")
                    continue
            
            print(f"\nFinal masked text: '{masked_text}'")
            return masked_text
            
        except Exception as e:
            print(f"Error in mask_entities: {e}")
            print(f"Text: {text}")
            print(f"Entities: {entities}")
            print(f"Style: {style}")
            return text  # Return original text on error

class TextDeidentificationFrame(tk.Frame):
    """Text de-identification panel with NLP-based PII detection"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.parent = parent
        self.pii_detector = PIIDetector()
        self.current_entities = []
        
        if not DEPENDENCIES_LOADED:
            self.show_dependency_error()
        else:
            self.setup_ui()
            
    def show_dependency_error(self):
        """Show error message when dependencies are missing"""
        message_frame = tk.Frame(self, bg="#ffffff", pady=20)
        message_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(message_frame,
                text="⚠️ PII Detection Feature Unavailable",
                font=("Segoe UI", 14, "bold"),
                fg="#f44336",
                bg="#ffffff").pack(pady=10)
                
        tk.Label(message_frame,
                text="Required dependencies are not installed.\n" +
                     "Please install them to use this feature:",
                font=("Segoe UI", 10),
                fg="#424242",
                bg="#ffffff").pack(pady=5)
                
        packages = ["presidio-analyzer", "presidio-anonymizer", 
                   "spacy", "flair", "transformers", "torch"]
        for pkg in packages:
            tk.Label(message_frame,
                    text=f"• {pkg}",
                    font=("Segoe UI", 9),
                    fg="#757575",
                    bg="#ffffff").pack(pady=2)
                    
        tk.Button(message_frame,
                 text="Install Dependencies",
                 command=lambda: check_dependencies(),
                 bg="#2196f3",
                 fg="white",
                 font=("Segoe UI", 9, "bold"),
                 relief=tk.FLAT,
                 padx=20,
                 pady=6).pack(pady=20)
        
    def setup_ui(self):
        """Setup the text de-identification UI"""
        # Main container with same style as parent
        main_container = tk.Frame(self, bg="#ffffff")
        main_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        # Top configuration panel
        config_frame = tk.Frame(main_container, bg="#ffffff", relief=tk.RAISED, bd=1)
        config_frame.pack(fill=tk.X, pady=4)
        
        # Model selection
        model_frame = tk.Frame(config_frame, bg="#ffffff")
        model_frame.pack(fill=tk.X, padx=8, pady=4)
        
        tk.Label(model_frame, text="NER Model:", bg="#ffffff", 
                font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=4)
                
        self.model_combo = ttk.Combobox(model_frame, 
                                       values=list(NLP_MODELS.keys()),
                                       width=30)
        self.model_combo.pack(side=tk.LEFT, padx=4)
        self.model_combo.set("flair/ner-english-large")  # Default model
        self.model_combo.bind("<<ComboboxSelected>>", self.on_model_change)
        
        # Entity type selection
        types_frame = tk.Frame(config_frame, bg="#ffffff")
        types_frame.pack(fill=tk.X, padx=8, pady=4)
        
        tk.Label(types_frame, text="Entity Types:", bg="#ffffff", 
                font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=4)
        
        # Create scrollable frame for entity types
        types_scroll = tk.Frame(types_frame, bg="#ffffff")
        types_scroll.pack(fill=tk.X, padx=4, pady=2)
        
        self.type_vars = {}
        row = 0
        col = 0
        for entity_type, info in PII_ENTITY_TYPES.items():
            var = tk.BooleanVar(value=True)
            self.type_vars[entity_type] = var
            
            cb = ttk.Checkbutton(types_scroll, text=entity_type, 
                                variable=var, command=self.on_type_change)
            cb.grid(row=row, column=col, sticky="w", padx=4, pady=2)
            
            col += 1
            if col > 3:  # 4 checkboxes per row
                col = 0
                row += 1
        
        # Select All/None buttons
        select_frame = tk.Frame(types_frame, bg="#ffffff")
        select_frame.pack(fill=tk.X, padx=4, pady=2)
        
        tk.Button(select_frame, text="Select All", command=self.select_all_types,
                  bg="#4caf50", fg="white", font=("Segoe UI", 8, "bold"),
                  relief=tk.FLAT, padx=8, pady=2).pack(side=tk.LEFT, padx=2)
                  
        tk.Button(select_frame, text="Select None", command=self.select_no_types,
                  bg="#f44336", fg="white", font=("Segoe UI", 8, "bold"),
                  relief=tk.FLAT, padx=8, pady=2).pack(side=tk.LEFT, padx=2)
        
        # Masking style
        style_frame = tk.Frame(config_frame, bg="#ffffff")
        style_frame.pack(fill=tk.X, padx=8, pady=4)
        
        tk.Label(style_frame, text="Masking Style:", bg="#ffffff",
                font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=4)
                
        self.style_combo = ttk.Combobox(style_frame,
                                       values=list(MASKING_STYLES.keys()),
                                       width=20)
        self.style_combo.pack(side=tk.LEFT, padx=4)
        self.style_combo.set("type_label")
        
        # Threshold slider
        threshold_frame = tk.Frame(config_frame, bg="#ffffff")
        threshold_frame.pack(fill=tk.X, padx=8, pady=4)
        
        tk.Label(threshold_frame, text="Confidence Threshold:", bg="#ffffff",
                font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=4)
                
        self.threshold_var = tk.DoubleVar(value=0.35)
        threshold_slider = ttk.Scale(threshold_frame, from_=0.0, to=1.0,
                                   orient="horizontal", length=200,
                                   variable=self.threshold_var,
                                   command=self.on_threshold_change)
        threshold_slider.pack(side=tk.LEFT, padx=4)
        
        self.threshold_label = tk.Label(threshold_frame, text="0.35",
                                      bg="#ffffff", width=4)
        self.threshold_label.pack(side=tk.LEFT, padx=4)
        
        # Split view for input/output
        paned = tk.PanedWindow(main_container, orient=tk.HORIZONTAL,
                              sashwidth=8, bg="#2196f3")
        paned.pack(fill=tk.BOTH, expand=True, pady=4)
        
        # Input panel
        input_frame = tk.Frame(paned, bg="#ffffff")
        paned.add(input_frame)
        
        tk.Label(input_frame, text="Input Text:", bg="#ffffff",
                font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=4, pady=2)
                
        self.input_text = scrolledtext.ScrolledText(input_frame, wrap=tk.WORD,
                                                   width=40, height=20)
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)
        
        # Button frame
        btn_frame = tk.Frame(input_frame, bg="#ffffff")
        btn_frame.pack(fill=tk.X, padx=4, pady=4)
        
        tk.Button(btn_frame, text="Detect & Mask", command=self.detect_and_mask,
                  bg="#2196f3", fg="white", font=("Segoe UI", 9, "bold"),
                  relief=tk.FLAT, padx=20, pady=6).pack(side=tk.LEFT, padx=2)
                  
        tk.Button(btn_frame, text="Clear", command=self.clear_text,
                  bg="#757575", fg="white", font=("Segoe UI", 9, "bold"),
                  relief=tk.FLAT, padx=20, pady=6).pack(side=tk.LEFT, padx=2)
        
        # Output panel
        output_frame = tk.Frame(paned, bg="#ffffff")
        paned.add(output_frame)
        
        tk.Label(output_frame, text="De-identified Text:", bg="#ffffff",
                font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=4, pady=2)
                
        self.output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD,
                                                    width=40, height=20)
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)
        
        # Status frame
        status_frame = tk.Frame(output_frame, bg="#ffffff")
        status_frame.pack(fill=tk.X, padx=4, pady=4)
        
        self.status_label = tk.Label(status_frame, text="Ready",
                                    bg="#ffffff", fg="#757575")
        self.status_label.pack(side=tk.LEFT)
        
        # Initialize the model
        self.initialize_default_model()
        
    def initialize_default_model(self):
        """Initialize the default NLP model"""
        model_name = self.model_combo.get()
        if self.pii_detector.initialize_models(model_name):
            self.status_label.config(text=f"Initialized {model_name}")
        else:
            self.status_label.config(text="Error initializing model")
            
    def on_model_change(self, event=None):
        """Handle model selection change"""
        model_name = self.model_combo.get()
        
        # Show loading indicator
        self.status_label.config(text=f"Loading {model_name}...")
        loading_window = tk.Toplevel(self.parent)
        loading_window.title("Loading Model")
        loading_window.geometry("300x100")
        loading_window.transient(self.parent)
        loading_window.grab_set()
        
        # Center on parent
        x = self.parent.winfo_x() + (self.parent.winfo_width() - 300) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - 100) // 2
        loading_window.geometry(f"+{x}+{y}")
        
        # Loading message
        tk.Label(loading_window, text=f"Loading {model_name}...",
                font=("Segoe UI", 10)).pack(pady=10)
        
        # Progress bar
        progress = ttk.Progressbar(loading_window, mode='indeterminate')
        progress.pack(fill=tk.X, padx=20, pady=10)
        progress.start()
        
        def load_model():
            try:
                if self.pii_detector.initialize_models(model_name):
                    self.status_label.config(text=f"Switched to {model_name}")
                else:
                    self.status_label.config(text="Error loading model")
                    messagebox.showerror("Error", f"Failed to load {model_name}")
            except Exception as e:
                self.status_label.config(text="Error loading model")
                messagebox.showerror("Error", f"Failed to load model: {str(e)}")
            finally:
                loading_window.destroy()
        
        # Start loading in thread
        threading.Thread(target=load_model, daemon=True).start()
            
    def on_type_change(self):
        """Handle entity type selection change"""
        if self.current_entities:
            self.update_output()
            
    def on_threshold_change(self, value):
        """Handle threshold slider change"""
        self.threshold_var.set(float(value))
        self.threshold_label.config(text=f"{float(value):.2f}")
        self.pii_detector.threshold = float(value)
        
        if self.current_entities:
            self.update_output()
            
    def select_all_types(self):
        """Select all entity types"""
        for var in self.type_vars.values():
            var.set(True)
        if self.current_entities:
            self.update_output()
            
    def select_no_types(self):
        """Deselect all entity types"""
        for var in self.type_vars.values():
            var.set(False)
        if self.current_entities:
            self.update_output()
            
    def get_selected_types(self):
        """Get list of selected entity types"""
        return [etype for etype, var in self.type_vars.items() if var.get()]
        
    def detect_and_mask(self):
        """Detect and mask PII in input text"""
        input_text = self.input_text.get("1.0", tk.END).strip()
        if not input_text:
            self.status_label.config(text="No input text")
            return
            
        self.status_label.config(text="Detecting entities...")
        self.update()

        # First check if model is initialized
        try:
            model_name = self.model_combo.get()
            if not model_name:
                messagebox.showwarning("Warning", "Please select a model first.")
                return

            if not self.pii_detector.is_initialized:
                if not self.pii_detector.initialize_models(model_name):
                    messagebox.showerror("Error", "Failed to initialize model. Please try again.")
                    return
        except Exception as e:
            messagebox.showerror("Error", f"Model initialization failed: {str(e)}")
            return
        
        # Run detection in a thread to avoid freezing UI
        def detect():
            try:
                selected_types = self.get_selected_types()
                if not selected_types:
                    self.after_idle(lambda: messagebox.showwarning("Warning", 
                        "Please select at least one entity type to detect."))
                    return

                # Print debug info
                print(f"\nStarting detection with model: {self.model_combo.get()}")
                print(f"Input text: '{input_text}'")
                print(f"Selected types: {selected_types}")
                print(f"Threshold: {self.threshold_var.get()}")

                self.current_entities = self.pii_detector.detect_entities(
                    input_text, selected_types)
                
                print(f"Detected {len(self.current_entities)} entities")
                for entity in self.current_entities:
                    print(f"Entity found: {entity}")
                
                self.after(100, self.update_output)
            except Exception as e:
                print(f"Error in detection thread: {e}")
                self.after_idle(lambda: messagebox.showerror("Error", 
                    f"Detection failed: {str(e)}"))
            
        threading.Thread(target=detect, daemon=True).start()
        
    def update_output(self):
        """Update output text with masked entities"""
        input_text = self.input_text.get("1.0", tk.END).strip()
        
        print("\nUpdating output with masked text")
        print(f"Original text: '{input_text}'")

        # Safety check
        if not hasattr(self, 'current_entities') or self.current_entities is None:
            print("No entities detected yet")
            self.status_label.config(text="No entities detected")
            return

        # Filter entities by selected types and threshold
        selected_types = self.get_selected_types()
        print(f"Selected types: {selected_types}")
        print(f"Current threshold: {self.threshold_var.get()}")

        filtered_entities = []
        try:
            filtered_entities = [ent for ent in self.current_entities
                               if ent["entity"] in selected_types
                               and ent["score"] >= self.threshold_var.get()]
            print(f"\nFiltered {len(filtered_entities)} entities from {len(self.current_entities)} total")
            for ent in filtered_entities:
                print(f"Including entity: {ent}")
        except Exception as e:
            print(f"Error filtering entities: {e}")
            self.status_label.config(text=f"Error filtering entities: {e}")
            return

        # Apply masking with error handling
        try:
            style = self.style_combo.get()
            print(f"\nApplying masking style: {style}")
            masked_text = self.pii_detector.mask_entities(input_text, filtered_entities, style)
            print(f"Masked text: '{masked_text}'")
        except Exception as e:
            print(f"Error masking entities: {e}")
            self.status_label.config(text=f"Error masking entities: {e}")
            return
        
        # Update output and status
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", masked_text)
        
        # Show detailed status
        if len(filtered_entities) > 0:
            details = "\n".join([
                f"• {ent['text']} ({ent['entity']}, score: {ent['score']:.2f})"
                for ent in filtered_entities
            ])
            status = (f"Found {len(filtered_entities)} entities of {len(self.current_entities)} detected:\n"
                    f"{details}")
            print("\nDetected Entities:")
            print(details)
        else:
            if len(self.current_entities) > 0:
                status = f"No entities matched the selected types/threshold (from {len(self.current_entities)} detected)"
                print("\nNo entities matched filters")
            else:
                status = "No entities detected in the text"
                print("\nNo entities found")
        
        self.status_label.config(text=status)
            
    def clear_text(self):
        """Clear input and output text"""
        self.input_text.delete("1.0", tk.END)
        self.output_text.delete("1.0", tk.END)
        self.current_entities = []
        self.status_label.config(text="Ready")

    def load_model_async(self, model_name):
        """Load NLP model asynchronously"""
        def load():
            try:
                self.status_text.set(f"Loading {model_name}...")
                self.progress_var.set(0)
                self.window.update()
                
                success = self.pii_detector.initialize_models(model_name)
                
                if success:
                    self.status_text.set(f"Loaded {model_name}")
                else:
                    self.status_text.set(f"Error loading {model_name}")
                    messagebox.showerror("Error", f"Failed to load {model_name}")
                    
                self.progress_var.set(100)
                
            except Exception as e:
                self.status_text.set("Error loading model")
                messagebox.showerror("Error", f"Failed to load model: {str(e)}")
                
        threading.Thread(target=load, daemon=True).start()

class UltraPrecisionRedactor:
    """Ultra-precise redaction engine with perfect boundary detection"""
    
    @staticmethod
    def find_ultra_precise_matches(page, regex_pattern):
        """Ultra-precise text matching using word-level data - FIXED SYNTAX"""
        matches = []
        
        try:
            # Get word-level data for precise boundaries
            words = page.get_text("words")
            full_text = page.get_text()
        
            # Find regex matches in full text
            for regex_match in re.finditer(regex_pattern, full_text, re.MULTILINE | re.IGNORECASE):
                matched_text = regex_match.group().strip()
        
                if matched_text and len(matched_text) > 0:
                    # For single character matches, use character-level precision
                    if len(matched_text) == 1:
                        char_rects = UltraPrecisionRedactor.get_character_level_boundaries(
                            words, matched_text, regex_match.start(), regex_match.end()
                        )
                        for rect in char_rects:
                            matches.append({
                                'rect': rect,
                                'text': matched_text
                            })
                    else:
                        # Find the exact word boundaries using word-level data
                        precise_rects = UltraPrecisionRedactor.get_word_level_boundaries(
                            words, matched_text
                        )
                    
                        for rect in precise_rects:
                            matches.append({
                                'rect': rect,
                                'text': matched_text
                            })
                    
        except Exception as e:
            print(f"Error in ultra-precise matching: {e}")
        
        return matches

    @staticmethod
    def get_word_level_boundaries(words, target_text):
        """Get precise boundaries using word-level data"""
        precise_rects = []
    
        try:
            target_lower = target_text.lower()
        
            # Method 1: Exact word match
            for word_data in words:
                x0, y0, x1, y1, word_text, block_no, line_no, word_no = word_data
            
                if word_text.lower() == target_lower:
                    precise_rects.append(fitz.Rect(x0, y0, x1, y1))
        
            if precise_rects:
                return precise_rects
        
            # Method 2: Partial match within words
            for word_data in words:
                x0, y0, x1, y1, word_text, block_no, line_no, word_no = word_data
            
                if target_lower in word_text.lower():
                    # Calculate precise position within the word
                    word_lower = word_text.lower()
                    start_pos = word_lower.find(target_lower)
                
                    if start_pos >= 0:
                        word_width = x1 - x0
                        char_width = word_width / len(word_text) if len(word_text) > 0 else word_width
                    
                        start_x = x0 + (start_pos * char_width)
                        end_x = start_x + (len(target_text) * char_width)
                    
                        precise_rects.append(fitz.Rect(start_x, y0, end_x, y1))
        
            if precise_rects:
                return precise_rects
        
            # Method 3: Multi-word match
            text_parts = target_text.split()
            if len(text_parts) > 1:
                # Find consecutive words that match
                for i in range(len(words) - len(text_parts) + 1):
                    word_sequence = []
                    for j in range(len(text_parts)):
                        if i + j < len(words):
                            word_sequence.append(words[i + j][4].lower())  # word_text
                
                    if ' '.join(word_sequence) == target_lower:
                        # Create bounding rectangle for all words
                        first_word = words[i]
                        last_word = words[i + len(text_parts) - 1]
                    
                        combined_rect = fitz.Rect(
                            first_word[0],  # x0 of first word
                            min(first_word[1], last_word[1]),  # min y0
                            last_word[2],   # x1 of last word
                            max(first_word[3], last_word[3])   # max y1
                        )
                    
                        precise_rects.append(combined_rect)
        
        except Exception as e:
            print(f"Error in word-level boundaries: {e}")
    
        return precise_rects
    
    @staticmethod
    def get_character_level_boundaries(words, target_char, match_start, match_end):
        """Get precise boundaries for single character matches"""
        precise_rects = []
    
        try:
            target_lower = target_char.lower()
        
            # Find all occurrences of the character in words
            for word_data in words:
                x0, y0, x1, y1, word_text, block_no, line_no, word_no = word_data
            
                # Check each character in the word
                for i, char in enumerate(word_text):
                    if char.lower() == target_lower:
                        # Calculate character position within word
                        word_width = x1 - x0
                        char_width = word_width / len(word_text) if len(word_text) > 0 else word_width
                        
                        char_x0 = x0 + (i * char_width)
                        char_x1 = char_x0 + char_width
                        
                        precise_rects.append(fitz.Rect(char_x0, y0, char_x1, y1))
        
        except Exception as e:
            print(f"Error in character-level boundaries: {e}")
    
        return precise_rects
    
    @staticmethod
    def apply_ultra_precise_redaction(page, matches):
        """Apply redaction with ultra-precise boundaries"""
        redaction_count = 0
        
        try:
            # First collect all redactions
            redactions = []
            for match in matches:
                try:
                    precise_rect = match['rect']
                    
                    # Validate and adjust rectangle if needed
                    if precise_rect.width > 0.1 and precise_rect.height > 0.1:
                        # Add some padding around the text
                        padding = 2
                        precise_rect.x0 -= padding
                        precise_rect.x1 += padding
                        precise_rect.y0 -= padding
                        precise_rect.y1 += padding
                        
                        redactions.append(precise_rect)
                        redaction_count += 1
                except Exception as e:
                    print(f"Error processing redaction: {e}")
                    continue
            
            # Merge overlapping redactions
            merged_redactions = []
            for rect in sorted(redactions, key=lambda r: (r.y0, r.x0)):
                merged = False
                for merged_rect in merged_redactions:
                    if merged_rect.intersects(rect):
                        merged_rect.include_rect(rect)
                        merged = True
                        break
                if not merged:
                    merged_redactions.append(rect)
            
            # Apply merged redactions
            for rect in merged_redactions:
                try:
                    # Add redaction annotation
                    annot = page.add_redact_annot(rect)
                    
                    # Set properties for better visibility
                    annot.set_fill_color((0, 0, 0))  # Black fill
                    annot.set_border_color((1, 0, 0))  # Red border
                    annot.set_border_width(0)  # No border
                    annot.update()
                    
                except Exception as e:
                    print(f"Error applying redaction: {e}")
                    continue
                    
            # Apply all redactions at once
            page.apply_redactions()
            
        except Exception as e:
            print(f"Error in redaction process: {e}")
            
        return redaction_count

class SmartRedactorEnhanced:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Smart PDF Redactor - Enhanced Text Selection & PII Detection - BY CA. Vaibhav Balar")
        
        # Check dependencies first
        if not check_dependencies():
            self.window.destroy()
            return
            
        # Hide console window for EXE
        if getattr(sys, 'frozen', False):
            self.window.withdraw()
            self.window.deiconify()
        
        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate window size
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        
        self.window.geometry(f"{window_width}x{window_height}")
        self.window.configure(bg="#f8f9fa")
        
        # Initialize PII detector
        self.pii_detector = PIIDetector()
        self.use_pii_detection = tk.BooleanVar(value=False)
        
        # Show startup message
        self.startup_message = messagebox.showinfo(
            "PII Detection Ready",
            "PII Detection feature is now available!\n\n" +
            "You can now:\n" +
            "1. Use the Text De-identification tab for text analysis\n" +
            "2. Enable smart PII detection in PDF redaction\n" +
            "3. Choose from multiple NLP models\n\n" +
            "Start by selecting a model and the types of data to detect."
        )
        
        # Add global Enter key binding
        self.window.bind('<Return>', self.on_global_enter)
        self.window.bind('<KP_Enter>', self.on_global_enter)
        
        # Core variables
        self.patterns = []
        self.pdf_path = None
        self.zoom = 1.2
        self.current_page = 0
        self.total_pages = 0
        self.canvas_img = None
        self.match_count = tk.StringVar()
        self.redaction_count = 0
        
        # UI variables
        self.status_text = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.current_inputs = []
        
        # Hit analysis variables
        self.total_hits = 0
        self.hit_details = {}
        self.hit_analysis_done = False
        
        # Enhanced font configuration
        self.logo_font = ("Segoe UI", 18, "bold")
        self.subtitle_font = ("Segoe UI", 10)
        self.section_font = ("Segoe UI", 11, "bold")
        self.label_font = ("Segoe UI", 9, "bold")
        self.text_font = ("Segoe UI", 9)
        self.small_font = ("Segoe UI", 8)
        self.button_font = ("Segoe UI", 9, "bold")
        
        # Initialize ultra-precision redactor
        self.redactor = UltraPrecisionRedactor()
        
        # Initialize text selection handler
        self.text_selector = TextSelectionHandler(self)
        
        self.setup_enhanced_ui()
        
    def on_global_enter(self, event):
        """Handle global Enter key press"""
        focused = self.window.focus_get()
        
        if isinstance(focused, tk.Button):
            focused.invoke()
            return 'break'
        
        if isinstance(focused, (ttk.Combobox, AutocompleteCombobox)):
            if hasattr(focused, 'on_enter'):
                focused.on_enter(event)
            return 'break'
            
        if isinstance(focused, tk.Entry):
            parent = focused.master
            for child in parent.winfo_children():
                if isinstance(child, tk.Button) and ("Add" in child.cget('text') or "Test" in child.cget('text')):
                    child.invoke()
                    return 'break'
        
        return None
        
    def setup_enhanced_ui(self):
        # Main container
        main_container = tk.Frame(self.window, bg="#f8f9fa")
        main_container.pack(fill=tk.BOTH, expand=True)

        # Notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        
        # PDF Redaction tab
        self.pdf_tab = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.pdf_tab, text=" 📄 PDF Redaction ")
        
        # Text De-identification tab
        self.text_tab = TextDeidentificationFrame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.text_tab, text=" 📝 Text De-identification ")
        
        # Pack notebook after adding tabs
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        
        # HEADER with enhanced branding (in PDF tab)
        header_frame = tk.Frame(self.pdf_tab, bg="#1a237e", height=100)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header_frame, bg="#1a237e")
        header_content.pack(expand=True, fill=tk.BOTH, padx=20, pady=8)
        
        # Logo section
        logo_section = tk.Frame(header_content, bg="#1a237e")
        logo_section.pack(side=tk.LEFT, fill=tk.Y)
        
        logo_frame = tk.Frame(logo_section, bg="#1a237e")
        logo_frame.pack(anchor="w", pady=8)
        
        # Selection icon
        perf_label = tk.Label(logo_frame, text="📝", font=("Segoe UI", 24), 
                             bg="#1a237e", fg="white")
        perf_label.pack(side=tk.LEFT, padx=(0, 12))
        
        # Title with selection branding
        title_frame = tk.Frame(logo_frame, bg="#1a237e")
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        title_label = tk.Label(title_frame, text="Smart PDF Redactor", 
                              font=self.logo_font, fg="white", bg="#1a237e")
        title_label.pack(anchor="w")
        
        # Selection branding - FIXED VISIBILITY WITH PROPER LAYOUT
        subtitle_label = tk.Label(title_frame, text="BY CA. Vaibhav Balar", 
                         font=("Segoe UI", 11, "bold"), fg="#4caf50", bg="#1a237e")
        subtitle_label.pack(anchor="w", pady=(4, 0))  # Added top padding, removed bottom padding
        
        # Selection badge
        version_frame = tk.Frame(header_content, bg="#1a237e")
        version_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        version_badge = tk.Label(version_frame, text="v5.0", 
                                font=("Segoe UI", 9, "bold"), 
                                bg="#ff9800", fg="white", 
                                padx=8, pady=4, relief=tk.FLAT)
        version_badge.pack(anchor="e", pady=2)
        
        # Separator
        separator_frame = tk.Frame(self.window, height=4, bg="#e0e0e0")
        separator_frame.pack(fill=tk.X)
        separator_frame.pack_propagate(False)
        
        # Main container
        main_container = tk.Frame(self.window, bg="#f8f9fa")
        main_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        # Paned window
        self.paned_window = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, 
                                          sashrelief=tk.RAISED, sashwidth=8, bg='#2196f3')
        self.paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Sidebar
        self.sidebar_frame = tk.Frame(self.paned_window, bg="#ffffff", bd=2, relief=tk.RAISED)
        self.paned_window.add(self.sidebar_frame, minsize=380, width=400)
        
        # PDF viewer panel
        self.pdf_frame = tk.Frame(self.paned_window, bg="#ffffff", bd=2, relief=tk.RAISED)
        self.paned_window.add(self.pdf_frame, minsize=500)
        
        self.setup_enhanced_sidebar()
        self.setup_enhanced_pdf_viewer()
        
        # Status bar
        status_frame = tk.Frame(self.window, bg="#37474f", height=36)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        status_container = tk.Frame(status_frame, bg="#37474f")
        status_container.pack(fill=tk.X, padx=12, pady=8)
        
        # Status icon
        status_icon = tk.Label(status_container, text="📝", font=("Segoe UI", 10), 
                              fg="#ff5722", bg="#37474f")
        status_icon.pack(side=tk.LEFT, padx=(0, 8))
        
        tk.Label(status_container, textvariable=self.status_text, 
                fg="white", bg="#37474f", font=self.text_font).pack(side=tk.LEFT)
        
        # Selection branding in status bar - CORRECTED
        brand_label = tk.Label(status_container, text="© CA. Vaibhav Balar • Smart Redaction", 
                              fg="#ff5722", bg="#37474f", font=("Segoe UI", 8, "bold"))
        brand_label.pack(side=tk.RIGHT, padx=(0, 12))
        
        # Progress bar
        progress_frame = tk.Frame(status_container, bg="#37474f")
        progress_frame.pack(side=tk.RIGHT, padx=8)
        
        self.mini_progress = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate', length=180)
        self.mini_progress.pack()
        
        self.status_text.set("📝 Ready - Select text from PDF preview to create patterns!")
        
    def setup_enhanced_sidebar(self):
        # Main scrollable container
        self.scrollable_sidebar = OptimizedScrollableFrame(self.sidebar_frame, bg="#ffffff")
        self.scrollable_sidebar.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        sidebar_content = self.scrollable_sidebar.inner_frame
        sidebar_content.configure(bg="#ffffff")
        
        # File Selection Section
        file_section_frame = tk.Frame(sidebar_content, bg="#ffffff", relief=tk.RAISED, bd=1)
        file_section_frame.pack(fill=tk.X, padx=4, pady=4)
        
        file_header = tk.Frame(file_section_frame, bg="#e8f5e8", height=30)
        file_header.pack(fill=tk.X)
        file_header.pack_propagate(False)
        
        tk.Label(file_header, text="📄 PDF Document", font=self.section_font, 
                bg="#e8f5e8", fg="#1a237e", pady=6).pack()
        
        file_content = tk.Frame(file_section_frame, bg="#ffffff")
        file_content.pack(fill=tk.X, padx=12, pady=8)
        
        file_btn = tk.Button(file_content, text="📂 Choose PDF File", command=self.select_pdf, 
                            bg="#4caf50", fg="white", font=self.button_font, 
                            relief=tk.FLAT, padx=20, pady=8, cursor="hand2",
                            activebackground="#388e3c", activeforeground="white")
        file_btn.pack(fill=tk.X, pady=4)
        
        self.file_label = tk.Label(file_content, text="No file selected", 
                                  bg="#ffffff", fg="#757575", font=self.text_font, 
                                  wraplength=340)
        self.file_label.pack(pady=4)
        
        # Text Selection Help Section
        help_section_frame = tk.Frame(sidebar_content, bg="#ffffff", relief=tk.RAISED, bd=1)
        help_section_frame.pack(fill=tk.X, padx=4, pady=4)
        
        help_header = tk.Frame(help_section_frame, bg="#fff3e0", height=30)
        help_header.pack(fill=tk.X)
        help_header.pack_propagate(False)
        
        tk.Label(help_header, text="📝 Text Selection Help", font=self.section_font, 
                bg="#fff3e0", fg="#e65100", pady=6).pack()
        
        help_content = tk.Frame(help_section_frame, bg="#ffffff")
        help_content.pack(fill=tk.X, padx=12, pady=8)
        
        help_text = """💡 How to use text selection:

• Double-click any word to select it
• Right-click to see pattern options
• Drag to select multiple words
• Smart suggestions based on text type
• Copy selected text to clipboard"""
        
        tk.Label(help_content, text=help_text, bg="#ffffff", fg="#424242", 
                font=self.small_font, justify=tk.LEFT, wraplength=320).pack(anchor="w")
        
        # Collapsible Sections
        
        # 1. Preset Patterns Section
        self.preset_section = CollapsibleFrame(sidebar_content, "Quick Patterns", "🎯")
        preset_content = self.preset_section.get_content_frame()
        
        tk.Label(preset_content, text="Select from predefined patterns:", 
                bg="#ffffff", font=self.text_font, fg="#424242").pack(anchor="w", padx=8, pady=(8,4))
        
        self.preset_combo = AutocompleteCombobox(preset_content, 
                                               values=list(PREDEFINED_PATTERNS.keys()), 
                                               font=self.text_font, height=8)
        self.preset_combo.pack(pady=4, fill=tk.X, padx=8)
        
        preset_btn = tk.Button(preset_content, text="➕ Add Quick Pattern", 
                              command=self.add_preset_pattern, bg="#4caf50", fg="white", 
                              font=self.button_font, relief=tk.FLAT, pady=6, cursor="hand2",
                              activebackground="#388e3c", activeforeground="white")
        preset_btn.pack(fill=tk.X, pady=8, padx=8)
        
        # 2. Custom Patterns Section
        self.custom_section = CollapsibleFrame(sidebar_content, "Custom Patterns", "⚙️")
        custom_content = self.custom_section.get_content_frame()
        
        tk.Label(custom_content, text="Create advanced pattern matching:", 
                bg="#ffffff", font=self.text_font, fg="#424242").pack(anchor="w", padx=8, pady=(8,4))
        
        self.match_type_combo = AutocompleteCombobox(custom_content, 
                                                   values=list(ADVANCED_MATCH_TYPES.keys()), 
                                                   font=self.text_font, height=12)  # Increased height to show more options
        self.match_type_combo.pack(pady=4, fill=tk.X, padx=8)
        self.match_type_combo.bind("<<ComboboxSelected>>", self.on_pattern_type_change)
        self.match_type_combo.bind("<KeyRelease>", self.on_pattern_type_change)
        
        # Description
        self.pattern_description = tk.Label(custom_content, text="", bg="#ffffff", 
                                          fg="#616161", font=self.small_font, 
                                          wraplength=340, justify=tk.LEFT)
        self.pattern_description.pack(anchor="w", pady=4, padx=8)
        
        # Dynamic input frame
        self.input_frame = tk.Frame(custom_content, bg="#ffffff")
        self.input_frame.pack(fill=tk.X, padx=8, pady=4)
        
        custom_btn = tk.Button(custom_content, text="➕ Add Custom Pattern", 
                              command=self.add_custom_pattern, bg="#ff9800", fg="white", 
                              font=self.button_font, relief=tk.FLAT, pady=6, cursor="hand2",
                              activebackground="#f57c00", activeforeground="white")
        custom_btn.pack(fill=tk.X, pady=8, padx=8)
        
        # 3. Active Patterns Section
        self.patterns_section = CollapsibleFrame(sidebar_content, "Active Patterns", "📋")
        patterns_content = self.patterns_section.get_content_frame()
        
        # Pattern count
        self.pattern_count_frame = tk.Frame(patterns_content, bg="#ffffff")
        self.pattern_count_frame.pack(fill=tk.X, padx=8, pady=4)
        
        self.pattern_count_label = tk.Label(self.pattern_count_frame, text="Patterns: 0", 
                                          bg="#ffffff", fg="#424242", font=self.text_font)
        self.pattern_count_label.pack(side=tk.LEFT)
        
        clear_all_btn = tk.Button(self.pattern_count_frame, text="🗑️ Clear All", 
                                 command=self.clear_all_patterns, bg="#f44336", fg="white", 
                                 font=("Segoe UI", 8, "bold"), relief=tk.FLAT, 
                                 padx=8, pady=2, cursor="hand2",
                                 activebackground="#d32f2f", activeforeground="white")
        clear_all_btn.pack(side=tk.RIGHT)
        
        # Patterns list with proper scrolling
        patterns_container = tk.Frame(patterns_content, bg="#ffffff")
        patterns_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # Create scrollable frame for patterns list
        self.patterns_scroll_frame = OptimizedScrollableFrame(patterns_container, bg="#ffffff")
        self.patterns_scroll_frame.pack(fill=tk.BOTH, expand=True)

        self.patterns_list_frame = self.patterns_scroll_frame.inner_frame
        
        # 4. Actions Section
        self.actions_section = CollapsibleFrame(sidebar_content, "Actions", "🚀")
        actions_content = self.actions_section.get_content_frame()
        
        # Preview button
        preview_btn = tk.Button(actions_content, text="👁️ Preview Matches", 
                               command=self.ultra_precise_preview, bg="#2196f3", fg="white", 
                               font=self.button_font, relief=tk.FLAT, pady=8, cursor="hand2",
                               activebackground="#1976d2", activeforeground="white")
        preview_btn.pack(fill=tk.X, pady=4, padx=8)
        
        # Redact button
        redact_btn = tk.Button(actions_content, text="🔒 Apply Redaction", 
                              command=self.apply_redaction, bg="#e91e63", fg="white", 
                              font=self.button_font, relief=tk.FLAT, pady=8, cursor="hand2",
                              activebackground="#c2185b", activeforeground="white")
        redact_btn.pack(fill=tk.X, pady=4, padx=8)
        
        # Match count display
        self.match_display = tk.Label(actions_content, textvariable=self.match_count, 
                                     bg="#ffffff", fg="#424242", font=self.text_font)
        self.match_display.pack(pady=8)
        
        # 5. Hit Analysis Section
        self.analysis_section = CollapsibleFrame(sidebar_content, "Hit Analysis", "📊")
        analysis_content = self.analysis_section.get_content_frame()
        
        analysis_btn = tk.Button(analysis_content, text="📊 Analyze Hits", 
                                command=self.analyze_hits, bg="#9c27b0", fg="white", 
                                font=self.button_font, relief=tk.FLAT, pady=8, cursor="hand2",
                                activebackground="#7b1fa2", activeforeground="white")
        analysis_btn.pack(fill=tk.X, pady=4, padx=8)
        
        self.analysis_display = tk.Label(analysis_content, text="No analysis performed yet", 
                                        bg="#ffffff", fg="#757575", font=self.small_font, 
                                        wraplength=320, justify=tk.LEFT)
        self.analysis_display.pack(pady=8, padx=8)
        
        # Update scroll region
        self.scrollable_sidebar.update_scroll_region()
        
    def setup_enhanced_pdf_viewer(self):
        # Ensure PDF frame takes up remaining space
        self.pdf_frame.pack_propagate(False)
        
        # PDF viewer header
        pdf_header = tk.Frame(self.pdf_frame, bg="#f5f5f5", height=50)
        pdf_header.pack(fill=tk.X)
        pdf_header.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(pdf_header, bg="#f5f5f5")
        header_content.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        # Title with selection icon
        title_frame = tk.Frame(header_content, bg="#f5f5f5")
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        tk.Label(title_frame, text="📝", font=("Segoe UI", 16), 
                bg="#f5f5f5", fg="#ff5722").pack(side=tk.LEFT, padx=(0, 8))
        
        tk.Label(title_frame, text="PDF Preview - Interactive Text Selection", 
                font=("Segoe UI", 12, "bold"), bg="#f5f5f5", fg="#1a237e").pack(side=tk.LEFT)
        
        # Controls
        controls_frame = tk.Frame(header_content, bg="#f5f5f5")
        controls_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Navigation
        nav_frame = tk.Frame(controls_frame, bg="#f5f5f5")
        nav_frame.pack(side=tk.RIGHT, padx=8)
        
        self.prev_btn = tk.Button(nav_frame, text="◀", command=self.prev_page, 
                                 bg="#2196f3", fg="white", font=("Segoe UI", 10, "bold"), 
                                 relief=tk.FLAT, padx=12, pady=4, cursor="hand2",
                                 activebackground="#1976d2", activeforeground="white")
        self.prev_btn.pack(side=tk.LEFT, padx=2)
        
        self.page_label = tk.Label(nav_frame, text="0 / 0", bg="#f5f5f5", 
                                  fg="#424242", font=("Segoe UI", 10, "bold"))
        self.page_label.pack(side=tk.LEFT, padx=8)
        
        self.next_btn = tk.Button(nav_frame, text="▶", command=self.next_page, 
                                 bg="#2196f3", fg="white", font=("Segoe UI", 10, "bold"), 
                                 relief=tk.FLAT, padx=12, pady=4, cursor="hand2",
                                 activebackground="#1976d2", activeforeground="white")
        self.next_btn.pack(side=tk.LEFT, padx=2)
        
        # Zoom controls
        zoom_frame = tk.Frame(controls_frame, bg="#f5f5f5")
        zoom_frame.pack(side=tk.RIGHT, padx=8)
        
        zoom_out_btn = tk.Button(zoom_frame, text="🔍-", command=self.zoom_out, 
                                bg="#757575", fg="white", font=("Segoe UI", 9, "bold"), 
                                relief=tk.FLAT, padx=8, pady=4, cursor="hand2",
                                activebackground="#616161", activeforeground="white")
        zoom_out_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_label = tk.Label(zoom_frame, text="120%", bg="#f5f5f5", 
                                  fg="#424242", font=("Segoe UI", 9, "bold"))
        self.zoom_label.pack(side=tk.LEFT, padx=4)
        
        zoom_in_btn = tk.Button(zoom_frame, text="🔍+", command=self.zoom_in, 
                               bg="#757575", fg="white", font=("Segoe UI", 9, "bold"), 
                               relief=tk.FLAT, padx=8, pady=4, cursor="hand2",
                               activebackground="#616161", activeforeground="white")
        zoom_in_btn.pack(side=tk.LEFT, padx=2)
        
        # PDF canvas container - FIXED OVERLAP ISSUE
        canvas_container = tk.Frame(self.pdf_frame, bg="#ffffff")
        canvas_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))
        
        # Selection instructions - MOVED TO TOP AND MADE COMPACT
        instructions_frame = tk.Frame(canvas_container, bg="#e3f2fd", relief=tk.FLAT, bd=1, height=25)
        instructions_frame.pack(fill=tk.X, pady=(0, 4))
        instructions_frame.pack_propagate(False)
        
        instructions_text = "💡 Double-click to select words • Right-click for pattern options"
        tk.Label(instructions_frame, text=instructions_text, bg="#e3f2fd", fg="#1565c0", 
                font=("Segoe UI", 8), pady=2).pack()
        
        # Canvas with scrollbars
        self.canvas = tk.Canvas(canvas_container, bg="#ffffff", highlightthickness=0)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_container, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack scrollbars and canvas
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Setup text selection bindings
        self.text_selector.setup_selection_bindings(self.canvas)
        
        # Mouse wheel scrolling
        self.canvas.bind("<MouseWheel>", self.on_canvas_scroll)
        
    def on_canvas_scroll(self, event):
        """Handle mouse wheel scrolling on canvas"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def on_pattern_type_change(self, event=None):
        """Handle pattern type selection change"""
        try:
            pattern_type = self.match_type_combo.get()
            
            if pattern_type in ADVANCED_MATCH_TYPES:
                pattern_info = ADVANCED_MATCH_TYPES[pattern_type]
                
                # Update description
                self.pattern_description.config(text=pattern_info["description"])
                
                # Clear existing inputs
                for widget in self.input_frame.winfo_children():
                    widget.destroy()
                self.current_inputs = []
                
                # Create input fields
                for i, input_label in enumerate(pattern_info["inputs"]):
                    label = tk.Label(self.input_frame, text=f"{input_label}:", 
                                   bg="#ffffff", font=self.label_font, fg="#424242")
                    label.pack(anchor="w", pady=(4, 2))
                    
                    entry = tk.Entry(self.input_frame, font=self.text_font, relief=tk.FLAT, 
                                   bd=1, bg="#f9f9f9", fg="#424242")
                    entry.pack(fill=tk.X, pady=(0, 4))
                    entry.bind('<Return>', lambda e: self.add_custom_pattern())
                    
                    self.current_inputs.append(entry)
                    
                    if i == 0:  # Focus first input
                        entry.focus_set()
            else:
                self.pattern_description.config(text="")
                for widget in self.input_frame.winfo_children():
                    widget.destroy()
                self.current_inputs = []
                
        except Exception as e:
            print(f"Error in pattern type change: {e}")
            
    def add_preset_pattern(self):
        """Add a preset pattern - FIXED VERSION"""
        try:
            pattern_name = self.preset_combo.get().strip()
            
            if not pattern_name:
                messagebox.showwarning("Warning", "Please select a pattern from the dropdown.")
                return
            
            if pattern_name not in PREDEFINED_PATTERNS:
                messagebox.showerror("Error", "Invalid pattern selected.")
                return
            
            # Check for duplicates
            for existing_pattern in self.patterns:
                if existing_pattern["label"] == pattern_name:
                    messagebox.showwarning("Warning", f"Pattern '{pattern_name}' already exists.")
                    return
        
            pattern = PREDEFINED_PATTERNS[pattern_name]
            color = self.generate_random_color()
        
            new_pattern = {
                "label": pattern_name,
                "regex": pattern,
                "color": color,
                "type": "preset"
            }
        
            self.patterns.append(new_pattern)
            self.add_pattern_display(color, pattern_name, len(self.patterns) - 1, pattern)
            self.update_pattern_count()
            self.preset_combo.set("")
            self.status_text.set(f"➕ Added preset pattern: {pattern_name}")
        
            # Auto-preview if PDF is loaded
            if self.pdf_path:
                self.ultra_precise_preview()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add preset pattern: {str(e)}")
            
    def add_custom_pattern(self):
        """Add a custom pattern"""
        try:
            pattern_type = self.match_type_combo.get().strip()
            
            if not pattern_type or pattern_type not in ADVANCED_MATCH_TYPES:
                messagebox.showwarning("Warning", "Please select a valid pattern type.")
                return
                
            pattern_info = ADVANCED_MATCH_TYPES[pattern_type]
            
            # Get input values
            input_values = []
            for entry in self.current_inputs:
                value = entry.get().strip()
                if not value:
                    messagebox.showwarning("Warning", "Please fill in all required fields.")
                    return
                input_values.append(value)
            
            # Generate regex
            if len(input_values) == 0:
                regex = pattern_info["generator"]()
            elif len(input_values) == 1:
                regex = pattern_info["generator"](input_values[0])
            elif len(input_values) == 2:
                regex = pattern_info["generator"](input_values[0], input_values[1])
            else:
                regex = pattern_info["generator"](*input_values)
            
            # Create label
            if input_values:
                label = f"{pattern_type}: {' + '.join(input_values)}"
            else:
                label = pattern_type
            
            color = self.generate_random_color()
            
            new_pattern = {
                "label": label,
                "regex": regex,
                "color": color,
                "type": "custom"
            }
            
            self.patterns.append(new_pattern)
            self.add_pattern_display(color, label, len(self.patterns) - 1, regex)
            
            self.update_pattern_count()
            
            # Clear inputs
            for entry in self.current_inputs:
                entry.delete(0, tk.END)
            self.match_type_combo.set("")
            self.pattern_description.config(text="")
            
            self.status_text.set(f"➕ Added custom pattern: {label}")
            
            # Auto-preview if PDF is loaded
            if self.pdf_path:
                self.ultra_precise_preview()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add custom pattern: {str(e)}")
            
    def add_pattern_display(self, color, label, index, regex):
        """Add pattern display widget"""
        try:
            pattern_frame = tk.Frame(self.patterns_list_frame, bg="#f9f9f9", 
                                   relief=tk.RAISED, bd=1)
            pattern_frame.pack(fill=tk.X, pady=2)
            
            # Color indicator
            color_frame = tk.Frame(pattern_frame, bg=color, width=4)
            color_frame.pack(side=tk.LEFT, fill=tk.Y)
            color_frame.pack_propagate(False)
            
            # Content frame
            content_frame = tk.Frame(pattern_frame, bg="#f9f9f9")
            content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=4)
            
            # Label
            label_text = label if len(label) <= 35 else label[:32] + "..."
            pattern_label = tk.Label(content_frame, text=label_text, bg="#f9f9f9", 
                                   fg="#424242", font=("Segoe UI", 9, "bold"), anchor="w")
            pattern_label.pack(anchor="w")
            
            # Regex (truncated)
            regex_text = regex if len(regex) <= 40 else regex[:37] + "..."
            regex_label = tk.Label(content_frame, text=f"Regex: {regex_text}", 
                                 bg="#f9f9f9", fg="#757575", font=("Segoe UI", 8), anchor="w")
            regex_label.pack(anchor="w")
            
            # Remove button
            remove_btn = tk.Button(pattern_frame, text="✕", command=lambda: self.remove_pattern(index), 
                                 bg="#f44336", fg="white", font=("Segoe UI", 8, "bold"), 
                                 relief=tk.FLAT, padx=6, pady=2, cursor="hand2",
                                 activebackground="#d32f2f", activeforeground="white")
            remove_btn.pack(side=tk.RIGHT, padx=4)
            
            # Store reference
            pattern_frame.pattern_index = index
            
            # Update scroll region after adding pattern
            self.patterns_scroll_frame.update_scroll_region()
            
        except Exception as e:
            print(f"Error adding pattern display: {e}")
            
    def remove_pattern(self, index):
        """Remove a pattern"""
        try:
            if 0 <= index < len(self.patterns):
                removed_pattern = self.patterns.pop(index)
                self.refresh_patterns_display()
                self.update_pattern_count()
                self.status_text.set(f"🗑️ Removed pattern: {removed_pattern['label']}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove pattern: {str(e)}")
            
    def clear_all_patterns(self):
        """Clear all patterns"""
        if self.patterns:
            if messagebox.askyesno("Confirm", "Are you sure you want to clear all patterns?"):
                self.patterns.clear()
                self.refresh_patterns_display()
                self.update_pattern_count()
                self.status_text.set("🗑️ All patterns cleared")
                
    def refresh_patterns_display(self):
        """Refresh the patterns display with proper scrolling"""
        try:
            # Clear existing displays
            for widget in self.patterns_list_frame.winfo_children():
                widget.destroy()
            
            # Re-add all patterns
            for i, pattern in enumerate(self.patterns):
                self.add_pattern_display(pattern["color"], pattern["label"], i, pattern["regex"])
        
            # Update scroll region for patterns
            self.patterns_scroll_frame.update_scroll_region()
            
        except Exception as e:
            print(f"Error refreshing patterns display: {e}")
            
    def update_pattern_count(self):
        """Update pattern count display"""
        count = len(self.patterns)
        self.pattern_count_label.config(text=f"Patterns: {count}")
        
    def generate_random_color(self):
        """Generate a random color for pattern highlighting"""
        colors = ["#ff5722", "#e91e63", "#9c27b0", "#673ab7", "#3f51b5", 
                 "#2196f3", "#03a9f4", "#00bcd4", "#009688", "#4caf50", 
                 "#8bc34a", "#cddc39", "#ffeb3b", "#ffc107", "#ff9800"]
        return random.choice(colors)
        
    def on_pii_detection_toggle(self):
        """Handle PII detection toggle"""
        if self.use_pii_detection.get():
            if not DEPENDENCIES_LOADED:
                messagebox.showwarning(
                    "Dependencies Required",
                    "PII detection requires additional packages.\n" +
                    "Please install them from the Text De-identification tab."
                )
                self.use_pii_detection.set(False)
                return
                
            if not self.pii_detector.ensure_initialized():
                messagebox.showwarning(
                    "Initialization Failed",
                    "Failed to initialize PII detection.\n" +
                    "Please check your installation and try again."
                )
                self.use_pii_detection.set(False)
                return
                
            self.load_model_async(self.pii_model_combo.get())
    
    def on_model_change(self, event=None):
        """Handle model selection change"""
        if self.use_pii_detection.get():
            self.load_model_async(self.pii_model_combo.get())
    
    def on_threshold_change(self, value):
        """Handle threshold slider change"""
        value = float(value)
        self.threshold_label.config(text=f"{value:.2f}")
        self.pii_detector.threshold = value
        
    def select_all_entities(self):
        """Select all entity types"""
        for var in self.pdf_type_vars.values():
            var.set(True)
    
    def select_no_entities(self):
        """Deselect all entity types"""
        for var in self.pdf_type_vars.values():
            var.set(False)
        
    def generate_entity_color(self, entity_type):
        """Generate consistent color for entity type"""
        # Define fixed colors for common entity types
        entity_colors = {
            "PERSON": "#e91e63",  # Pink
            "LOCATION": "#2196f3",  # Blue
            "ORGANIZATION": "#4caf50",  # Green
            "EMAIL": "#ff9800",  # Orange
            "PHONE_NUMBER": "#9c27b0",  # Purple
            "CREDIT_CARD": "#f44336",  # Red
            "ID_NUMBER": "#00bcd4",  # Cyan
            "DATE_TIME": "#ffc107",  # Amber
            "CRYPTO": "#ff5722",  # Deep Orange
            "BANK_ACCOUNT": "#673ab7",  # Deep Purple
            "US_SSN": "#e65100",  # Dark Orange
            "US_PASSPORT": "#1976d2",  # Dark Blue
            "MEDICAL_RECORD": "#c2185b"  # Dark Pink
        }
        
        # Return fixed color if defined, otherwise generate random color
        return entity_colors.get(entity_type, self.generate_random_color())
        
    def select_pdf(self):
        """Select PDF file"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            
            if file_path:
                self.pdf_path = file_path
                filename = os.path.basename(file_path)
                self.file_label.config(text=f"📄 {filename}", fg="#4caf50")
                
                # RESET PATTERNS WHEN NEW FILE IS LOADED
                if messagebox.askyesno("Reset Patterns", "Clear existing patterns for new file?"):
                    self.patterns.clear()
                    self.refresh_patterns_display()
                    self.update_pattern_count()
                    self.status_text.set("🗑️ Patterns cleared for new file")
            
                # Load PDF
                self.load_pdf()
                self.status_text.set(f"📄 Loaded: {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select PDF: {str(e)}")
            
    def load_pdf(self):
        """Load PDF and display first page"""
        try:
            if not self.pdf_path:
                return
                
            doc = fitz.open(self.pdf_path)
            self.total_pages = len(doc)
            self.current_page = 0
            doc.close()
            
            self.update_page_display()
            self.display_pdf_page()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
            
    def display_pdf_page(self):
        """Display current PDF page with enhanced rendering"""
        try:
            if not self.pdf_path:
                return
                
            doc = fitz.open(self.pdf_path)
            page = doc[self.current_page]
            
            # Render page with high quality
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # Convert to PIL Image
            img_data = pix.tobytes("ppm")
            pil_image = Image.open(io.BytesIO(img_data))
            
            # Convert to PhotoImage
            self.canvas_img = ImageTk.PhotoImage(pil_image)
            
            # Clear canvas and display image
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor="nw", image=self.canvas_img)
            
            # Update scroll region
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            doc.close()
            
        except Exception as e:
            print(f"Error displaying PDF page: {e}")
            
    def update_page_display(self):
        """Update page navigation display"""
        self.page_label.config(text=f"{self.current_page + 1} / {self.total_pages}")
        self.prev_btn.config(state="normal" if self.current_page > 0 else "disabled")
        self.next_btn.config(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
        
    def prev_page(self):
        """Go to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()
            self.display_pdf_page()
        
            # MAINTAIN PREVIEW WHEN CHANGING PAGES
            if self.patterns:
                self.ultra_precise_preview()

    def next_page(self):
        """Go to next page"""
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_page_display()
            self.display_pdf_page()
        
            # MAINTAIN PREVIEW WHEN CHANGING PAGES
            if self.patterns:
                self.ultra_precise_preview()
            
    def zoom_in(self):
        """Zoom in"""
        if self.zoom < 3.0:
            self.zoom += 0.2
            self.zoom_label.config(text=f"{int(self.zoom * 100)}%")
            self.display_pdf_page()
        
            # MAINTAIN PREVIEW WHEN ZOOMING
            if self.patterns:
                self.ultra_precise_preview()

    def zoom_out(self):
        """Zoom out"""
        if self.zoom > 0.4:
            self.zoom -= 0.2
            self.zoom_label.config(text=f"{int(self.zoom * 100)}%")
            self.display_pdf_page()
        
            # MAINTAIN PREVIEW WHEN ZOOMING
            if self.patterns:
                self.ultra_precise_preview()
            
    def ultra_precise_preview(self):
        """Ultra-precise preview with perfect boundary detection and PII detection"""
        try:
            if not self.pdf_path:
                messagebox.showwarning("Warning", "Please select a PDF file first.")
                return
                
            if not self.patterns and not self.use_pii_detection.get():
                messagebox.showwarning("Warning", "Please add patterns or enable PII detection.")
                return
                
            self.progress_var.set(0)
            self.status_text.set("🔍 Analyzing with ultra-precision...")
            self.window.update()
            
            doc = fitz.open(self.pdf_path)
            page = doc[self.current_page]
            
            # Clear previous highlights
            self.canvas.delete("highlight")
            
            total_matches = 0
            
            # Process patterns first
            pattern_progress = 50 if self.use_pii_detection.get() else 100
            for i, pattern in enumerate(self.patterns):
                try:
                    # Update progress
                    progress = (i / len(self.patterns)) * pattern_progress
                    self.progress_var.set(progress)
                    self.window.update()
                    
                    # Validate regex pattern before using
                    try:
                        re.compile(pattern["regex"])
                    except re.error as regex_error:
                        print(f"Invalid regex in pattern {pattern['label']}: {regex_error}")
                        continue
                    
                    # Find ultra-precise matches
                    matches = self.redactor.find_ultra_precise_matches(page, pattern["regex"])
                    
                    # Draw highlights on canvas
                    for match in matches:
                        rect = match['rect']
                        
                        # Convert PDF coordinates to canvas coordinates
                        canvas_x1 = rect.x0 * self.zoom
                        canvas_y1 = rect.y0 * self.zoom
                        canvas_x2 = rect.x1 * self.zoom
                        canvas_y2 = rect.y1 * self.zoom
                        
                        # Draw highlight rectangle
                        if canvas_x2 > canvas_x1 and canvas_y2 > canvas_y1:
                            self.canvas.create_rectangle(
                                canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                                outline=pattern["color"], width=2,
                                fill=pattern["color"], stipple="gray25",
                                tags="highlight"
                            )
                            total_matches += 1
                            
                except Exception as e:
                    print(f"Error processing pattern {pattern['label']}: {e}")
                    continue
                    
            # Process PII detection if enabled
            if self.use_pii_detection.get():
                try:
                    # Get page text
                    page_text = page.get_text()
                    
                    # Get selected entity types
                    selected_types = [etype for etype, var in self.pdf_type_vars.items()
                                    if var.get()]
                    
                    # Detect entities
                    self.progress_var.set(75)
                    self.window.update()
                    
                    entities = self.pii_detector.detect_entities(
                        page_text,
                        selected_types
                    )
                    
                    # Filter by threshold
                    threshold = self.pdf_threshold_var.get()
                    entities = [ent for ent in entities if ent["score"] >= threshold]
                    
                    # Draw highlights for entities
                    for entity in entities:
                        try:
                            # Find text position on page
                            text_instances = page.search_for(entity["text"])
                            
                            for inst in text_instances:
                                # Convert PDF coordinates to canvas coordinates
                                canvas_x1 = inst.x0 * self.zoom
                                canvas_y1 = inst.y0 * self.zoom
                                canvas_x2 = inst.x1 * self.zoom
                                canvas_y2 = inst.y1 * self.zoom
                                
                                # Generate consistent color for entity type
                                color = self.generate_entity_color(entity["entity"])
                                
                                # Draw highlight rectangle
                                if canvas_x2 > canvas_x1 and canvas_y2 > canvas_y1:
                                    self.canvas.create_rectangle(
                                        canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                                        outline=color, width=2,
                                        fill=color, stipple="gray25",
                                        tags="highlight"
                                    )
                                    total_matches += 1
                                    
                        except Exception as e:
                            print(f"Error highlighting entity: {e}")
                            continue
                            
                except Exception as e:
                    print(f"Error in PII detection: {e}")
                    
            # Process patterns first
            for i, pattern in enumerate(self.patterns):
                try:
                    # Update progress
                    progress = (i / len(self.patterns)) * pattern_progress
                    self.progress_var.set(progress)
                    self.window.update()
                    
                    # Validate regex pattern before using
                    try:
                        re.compile(pattern["regex"])
                    except re.error as regex_error:
                        print(f"Invalid regex in pattern {pattern['label']}: {regex_error}")
                        continue
                    
                    # Find ultra-precise matches
                    matches = self.redactor.find_ultra_precise_matches(page, pattern["regex"])
                    
                    # Draw highlights on canvas
                    for match in matches:
                        rect = match['rect']
                        
                        # Convert PDF coordinates to canvas coordinates
                        canvas_x1 = rect.x0 * self.zoom
                        canvas_y1 = rect.y0 * self.zoom
                        canvas_x2 = rect.x1 * self.zoom
                        canvas_y2 = rect.y1 * self.zoom
                        
                        # Ensure valid rectangle coordinates
                        if canvas_x2 > canvas_x1 and canvas_y2 > canvas_y1:
                            # Draw highlight rectangle
                            self.canvas.create_rectangle(
                                canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                                outline=pattern["color"], width=2, fill=pattern["color"],
                                stipple="gray25", tags="highlight"
                            )
                            
                            total_matches += 1
                            
                except Exception as e:
                    print(f"Error processing pattern {pattern['label']}: {e}")
                    continue
            
            doc.close()
            
            # Update display
            self.progress_var.set(100)
            self.match_count.set(f"📊 Found {total_matches} matches on page {self.current_page + 1}")
            self.status_text.set(f"✅ Ultra-precise preview complete - {total_matches} matches found")
            
        except Exception as e:
            messagebox.showerror("Error", f"Preview failed: {str(e)}")
            self.progress_var.set(0)
            
    def apply_redaction(self):
        """Apply ultra-precise redaction with PII detection"""
        try:
            if not self.pdf_path:
                messagebox.showwarning("Warning", "Please select a PDF file first.")
                return
                
            if not self.patterns and not self.use_pii_detection.get():
                messagebox.showwarning("Warning", "Please add patterns or enable PII detection.")
                return
                
            # Confirm redaction
            if not messagebox.askyesno("Confirm Redaction", 
                                     "This will permanently redact the PDF. Continue?"):
                return
                
            # Select output file
            output_path = filedialog.asksaveasfilename(
                title="Save Redacted PDF",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if not output_path:
                return
                
            self.progress_var.set(0)
            self.status_text.set("🔒 Applying ultra-precise redaction...")
            self.window.update()
            
            doc = fitz.open(self.pdf_path)
            total_redactions = 0
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Update progress
                base_progress = (page_num / len(doc)) * 90
                self.progress_var.set(base_progress)
                self.window.update()
                
                # Process patterns
                if self.patterns:
                    for pattern in self.patterns:
                        try:
                            # Find ultra-precise matches
                            matches = self.redactor.find_ultra_precise_matches(page, pattern["regex"])
                            
                            # Apply redaction
                            redaction_count = self.redactor.apply_ultra_precise_redaction(page, matches)
                            total_redactions += redaction_count
                            
                        except Exception as e:
                            print(f"Error redacting pattern {pattern['label']} on page {page_num + 1}: {e}")
                            continue
                
                # Process PII detection
                if self.use_pii_detection.get():
                    if not DEPENDENCIES_LOADED:
                        messagebox.showwarning("PII Detection Unavailable",
                            "PII detection requires additional dependencies.\n" +
                            "Please install them from the Text De-identification tab.")
                        continue
                        
                    if not self.pii_detector.ensure_initialized():
                        messagebox.showwarning("PII Detection Error",
                            "Failed to initialize PII detection.\n" +
                            "Continuing with pattern-based redaction only.")
                        continue
                        
                    try:
                        # Get page text
                        page_text = page.get_text()
                        
                        # Get selected entity types
                        selected_types = [etype for etype, var in self.pdf_type_vars.items()
                                        if var.get()]
                        
                        # Update progress
                        self.progress_var.set(base_progress + 5)
                        self.window.update()
                        
                        # Detect entities
                        entities = self.pii_detector.detect_entities(
                            page_text,
                            selected_types
                        )
                        
                        # Filter by threshold
                        threshold = self.pdf_threshold_var.get()
                        entities = [ent for ent in entities if ent["score"] >= threshold]
                        
                        # Apply redactions for entities
                        for entity in entities:
                            try:
                                # Find text position on page
                                text_instances = page.search_for(entity["text"])
                                
                                for inst in text_instances:
                                    # Create redaction annotation
                                    page.add_redact_annot(inst, fill=(0, 0, 0))
                                    total_redactions += 1
                                    
                            except Exception as e:
                                print(f"Error redacting entity: {e}")
                                continue
                                
                    except Exception as e:
                        print(f"Error in PII detection on page {page_num + 1}: {e}")
                
                # Apply redactions to page
                page.apply_redactions()
            
            # Save redacted PDF
            self.progress_var.set(95)
            self.status_text.set("💾 Saving redacted PDF...")
            self.window.update()
            
            # OLD: doc.save(output_path)
            # NEW: Use incremental save to preserve file size
            doc.save(output_path, garbage=4, deflate=True, clean=True)
            doc.close()
            
            self.progress_var.set(100)
            self.redaction_count = total_redactions
            
            # Success message
            success_msg = f"✅ Redaction Complete!\n\n"
            success_msg += f"📊 Total redactions: {total_redactions}\n"
            success_msg += f"📄 Pages processed: {self.total_pages}\n"
            success_msg += f"💾 Saved to: {os.path.basename(output_path)}"
            
            messagebox.showinfo("Success", success_msg)
            self.status_text.set(f"✅ Redaction complete - {total_redactions} items redacted")
            
            # Ask to open output folder
            if messagebox.askyesno("Open Folder", "Would you like to open the output folder?"):
                output_dir = os.path.dirname(output_path)
                if os.name == 'nt':  # Windows
                    os.startfile(output_dir)
                elif os.name == 'posix':  # macOS and Linux
                    os.system(f'open "{output_dir}"' if sys.platform == 'darwin' else f'xdg-open "{output_dir}"')
                    
        except Exception as e:
            messagebox.showerror("Error", f"Redaction failed: {str(e)}")
            self.progress_var.set(0)
            
    def analyze_hits(self):
        """Analyze pattern hits across all pages"""
        try:
            if not self.pdf_path:
                messagebox.showwarning("Warning", "Please select a PDF file first.")
                return
                
            if not self.patterns:
                messagebox.showwarning("Warning", "Please add at least one pattern.")
                return
                
            self.progress_var.set(0)
            self.status_text.set("📊 Analyzing hits across all pages...")
            self.window.update()
            
            doc = fitz.open(self.pdf_path)
            self.hit_details = {}
            self.total_hits = 0
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Update progress
                progress = (page_num / len(doc)) * 100
                self.progress_var.set(progress)
                self.window.update()
                
                for pattern in self.patterns:
                    pattern_label = pattern["label"]
                    
                    if pattern_label not in self.hit_details:
                        self.hit_details[pattern_label] = {"total": 0, "pages": []}
                    
                    try:
                        # Find matches on this page
                        matches = self.redactor.find_ultra_precise_matches(page, pattern["regex"])
                        page_hits = len(matches)
                        
                        if page_hits > 0:
                            self.hit_details[pattern_label]["total"] += page_hits
                            self.hit_details[pattern_label]["pages"].append({
                                "page": page_num + 1,
                                "hits": page_hits
                            })
                            self.total_hits += page_hits
                            
                    except Exception as e:
                        print(f"Error analyzing pattern {pattern_label} on page {page_num + 1}: {e}")
                        continue
            
            doc.close()
            
            # Display analysis results
            self.display_hit_analysis()
            self.hit_analysis_done = True
            
            self.progress_var.set(100)
            self.status_text.set(f"📊 Analysis complete - {self.total_hits} total hits found")
            
        except Exception as e:
            messagebox.showerror("Error", f"Hit analysis failed: {str(e)}")
            self.progress_var.set(0)
            
    def display_hit_analysis(self):
        """Display hit analysis results"""
        try:
            if not self.hit_details:
                self.analysis_display.config(text="No hits found in the document.")
                return
                
            analysis_text = f"📊 Total Hits: {self.total_hits}\n\n"
            
            for pattern_label, details in self.hit_details.items():
                total_hits = details["total"]
                pages_with_hits = len(details["pages"])
                
                analysis_text += f"🎯 {pattern_label}\n"
                analysis_text += f"   Hits: {total_hits} | Pages: {pages_with_hits}\n"
                
                # Show page details (limit to first 5 pages)
                page_details = details["pages"][:5]
                page_info = ", ".join([f"P{p['page']}({p['hits']})" for p in page_details])
                if len(details["pages"]) > 5:
                    page_info += f" +{len(details['pages']) - 5} more"
                    
                analysis_text += f"   Pages: {page_info}\n\n"
            
            self.analysis_display.config(text=analysis_text)
            
        except Exception as e:
            print(f"Error displaying hit analysis: {e}")
            
    def cleanup(self):
        """Clean up resources before exit"""
        try:
            # Clear any loaded models
            if hasattr(self, 'pii_detector'):
                self.pii_detector.models.clear()
            
            # Clear any temporary files
            import tempfile
            import shutil
            temp_dir = tempfile.gettempdir()
            pattern = os.path.join(temp_dir, "pii_detector_*")
            for temp_file in glob.glob(pattern):
                try:
                    if os.path.isfile(temp_file):
                        os.remove(temp_file)
                    elif os.path.isdir(temp_file):
                        shutil.rmtree(temp_file)
                except:
                    pass
        except:
            pass
    
    def run(self):
        """Run the application"""
        try:
            # Register cleanup
            self.window.protocol("WM_DELETE_WINDOW", 
                               lambda: (self.cleanup(), self.window.destroy()))
            
            # Start main loop
            self.window.mainloop()
            
        except KeyboardInterrupt:
            self.cleanup()
            self.window.quit()
        except Exception as e:
            print(f"Application error: {e}")
            messagebox.showerror("Application Error", 
                               f"An unexpected error occurred: {str(e)}")
            self.cleanup()

if __name__ == "__main__":
    # Enable debug output
    import sys
    def debug_print(*args, **kwargs):
        print(*args, file=sys.stderr, **kwargs)
        sys.stderr.flush()

    debug_print("Starting application...")
    
    try:
        # Set up better exception handling
        def show_error_dialog(title, message, error=None):
            try:
                import tkinter as tk
                from tkinter import messagebox
                
                # Create root window safely
                root = None
                try:
                    root = tk.Tk()
                    root.withdraw()
                    full_message = f"{message}\n\nError: {str(error)}" if error else message
                    messagebox.showerror(title, full_message)
                finally:
                    if root:
                        try:
                            root.destroy()
                        except tk.TclError:
                            pass
            except Exception as dialog_error:
                print(f"Failed to show error dialog: {dialog_error}")
                print(f"Original error: {error}")

        # Initialize application with proper error handling
        app = None
        try:
            app = SmartRedactorEnhanced()
            if hasattr(app, 'window') and app.window:
                try:
                    # Set up cleanup on window close
                    def on_closing():
                        try:
                            if app:
                                app.cleanup()
                            if hasattr(app, 'window') and app.window:
                                app.window.destroy()
                        except Exception as close_error:
                            print(f"Error during cleanup: {close_error}")
                    
                    app.window.protocol("WM_DELETE_WINDOW", on_closing)
                    app.run()
                except tk.TclError as tcl_error:
                    if "application has been destroyed" not in str(tcl_error):
                        raise
                except Exception as run_error:
                    raise run_error
        except Exception as e:
            show_error_dialog(
                "Application Error",
                "Failed to start the application.",
                e
            )
            if not getattr(sys, 'frozen', False):
                import traceback
                traceback.print_exc()
        finally:
            # Final cleanup
            if app:
                try:
                    app.cleanup()
                except:
                    pass
                    
    except Exception as e:
        print(f"Critical error: {e}")
        if not getattr(sys, 'frozen', False):
            import traceback
            traceback.print_exc()
