# Component-Based Authoring Application Level Add-In for Microsoft Word

## Introduction

The aim of this project is to allow component-based authoring in the word processor software [Microsoft Word][1]. It is
implemented as an [Application-Level Add-in][2] and is therefore
compatible with other Application-Level Add-ins and Document-Level Customizations.

## Requirements

### Usage

* [Microsoft Word][1] 2010
* [Microsoft Visual Studio Tools for Office Runtime (VSTO)][3] 2010
* (*optional*) [Microsoft Visual Studio Tools for Office Language Pack][4]2010

### Development

* [Microsoft Visual Studio][5] 2012 Professional Edition or above
* [MarkdownSharp][6] >= 1.13.0.0
* [NLog][7] >= 2.0.1.2
* [NLog Configuration][8] >= 2.0.1.2

## Features

* Completely unobtrusive:
  * Implemented as an [Application-Level Add-in][2], which can be used together with a [Document-Level Customization][9] and other Application-Level Add-ins.
  * Uses built-in Microsoft Word functions to implement the features.
  * Allows to enable and disable each automated function of the Add-in separately via a configuration dialog.
* Allows to create a [custom XML part][10] from a XML file in a subdirectory (default: `\XML`) of a "Microsoft Word" document.
  * A custom XML part is synchronized with the external XML file when a document is opened.
  * A custom XML part is identified via its default namespace (`xmlns` attribute of the root XML element). Therefore each XML file requires a unique default namespace declaration.
* Allows to create and bind [content controls][11] to a custom XML part via the graphical user interface. Two strategies are supported:
  * 1:1 mapping: Creates and binds one content control to one XML element or XML attribute.
  * List mapping: Creates and binds multiple content control to a XML element with child nodes. A content control is created for each XML element and XML attribute below the selected XML node.
* Allows to enforce copying of all styles from the attached Microsoft Word template into the active Microsoft Word document, overwriting any existing styles in the document that have the same name.
* Implements a simple integrated development environment that allows to interact with [fields][12] via the graphical user interface:
  * Create fields.
  * Format field results.
  * Perform actions with fields.
* Allows to compare a field result with the content of an included file.
* Allows to open referenced files via the graphical user interface.
* Allows to transfer content between the currently opened Microsoft Word document and a referenced file:
  * Overwrite the content of the referenced file with the content of the document.
  * Overwrite the content of the document with the content of the referenced file.
* Implements workarounds for known Microsoft Word software bugs:
  * Duplication of `IncludePicture` fields with Office Open XML (OOXML) file formats.
  * Relative path do not work with `IncludeText` and `IncludePicture` fields. Click [here][13] for a detailed discussion of the problem.
  * Fields in the header and the footer of a Microsoft Word document are not updated when the document is opened. Click [here][14] to read about the manual workaround.
  * Unable to deploy a Microsoft Word template file (with the DOTX or DOTM file extension) together with their associated Microsoft Word documents, since it is impossible to set a relative file path for the template file.

## Installation

Refer to the chapter *Installation* in the [user manual][15] (currently only available in the language German).

## Usage

Refer to the chapter *Functions* in the [user manual][15] (currently only available in the language German).

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://gnu.org/licenses/lgpl.txt>.

[1]: http://office.com/word
[2]: http://msdn.com/library/vstudio/bb386298
[3]: http://microsoft.com/download/details.aspx?id=39290
[4]: http://microsoft.com/download/details.aspx?id=39291
[5]: http://microsoft.com/visualstudio
[6]: http://nuget.org/packages/MarkdownSharp
[7]: http://nuget.org/packages/NLog
[8]: http://nuget.org/packages/NLog.Config
[9]: http://msdn.com/library/zcfbd2sk
[10]: http://msdn.com/library/bb608618
[11]: http://msdn.com/library/vstudio/bb157891
[12]: http://office.com/en-us/word-help/insert-and-format-field-codes-in-word-2010-HA101830917.aspx
[13]: http://word.mvps.org/FAQs/TblsFldsFms/includetextfieldscontent.htm#FilePaths
[14]: http://cybertext.wordpress.com/2008/01/25/update-fields-in-headers-and-footers
[15]: http://github.com/FlorianWolters/component-based-authoring-add-in-for-microsoft-word-docs/de/UserManual-DE.pdf
