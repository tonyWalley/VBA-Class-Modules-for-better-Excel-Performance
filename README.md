# VBA-Class-Modules-for-better-Excel-Performance

This repository contains a collection of VBA class modules designed to enhance Excel automation, streamline data processing, and improve code efficiency.

Class Modules Overview
FileMgr:

Provides utilities for file management, including path manipulation.
Methods to retrieve directory names, reduce folder paths, and create subfolders.
ListBoxTbl:

Manages ListBox controls in user forms.
Properties for ListBox names and item strings.
Methods for setting column values dynamically.
Logger:

Implements a logging system for debugging and tracking events.
Creates log files dynamically based on log types and timestamps.
Stores caller information and execution details.
RowDictionaryWrapper:

Handles row-level data storage using dictionaries.
Stores formulas and default values for table columns.
Flags last-item indicators for processing.
StringMgr:

Provides string manipulation functions.
Trims, pads, and ensures string length consistency.
Parses strings by delimiters and trims components.
tblDic:

Facilitates dictionary-based ListObject management.
Allows single-column updates and dynamic table processing.
Stores table structures and provides quick lookups.
WorksheetDictionary:

Maintains a dictionary of worksheets by their CodeName.
Enables quick retrieval of worksheet objects.
Helps manage large workbooks efficiently.
CodeSettings:

Centralized class for managing Excel settings.
Controls protection settings and sheet behavior.
Implements selection and accessibility restrictions.
CodeTable:

Provides a structured lookup system using dictionary-based tables.
Retrieves values based on key relationships.
Facilitates dynamic table referencing.
How to Use
Import the .cls files into your VBA project.
Use the provided classes to manage file operations, logging, list boxes, structured data storage, and advanced dictionary handling.
Installation
To use these modules:

Open the Visual Basic for Applications (VBA) Editor (ALT + F11 in Excel).
Go to File > Import File... and select the .cls files you need.
License
This project is released under the MIT License.

Contributions
Contributions and improvements are welcome! Feel free to submit a pull request or open an issue.
