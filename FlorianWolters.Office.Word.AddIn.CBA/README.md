# Component-Based Authoring Application Level Add-In for Microsoft Word

## Introduction

The aim of this project is to allow component-based authoring in the word processor software *Microsoft Word*. It is
implemented as an [Application-Level Add-in](http://msdn.microsoft.com/library/vstudio/bb386298.aspx) and is therefore
compatible with other Application-Level Add-ins and Document-Level Customizations.

## Requirements

### Usage

* [Microsoft Word](http://office.microsoft.com/word) 2010
* [Microsoft Visual Studio Tools for Office Runtime (VSTO)](http://microsoft.com/en-us/download/details.aspx?id=39290) 2010
* (*optional*) [Microsoft Visual Studio Tools for Office Language Pack](http://microsoft.com/en-us/download/details.aspx?id=39291) 2010

### Development

* [Microsoft Visual Studio](http://microsoft.com/visualstudio) 2012 Professional Edition or above
* [MarkdownSharp](http://nuget.org/packages/MarkdownSharp) >= 1.13.0.0
* [NLog](http://nuget.org/packages/NLog) >= 2.0.1.2
* [NLog Configuration](http://nuget.org/packages/NLog.Config) >= 2.0.1.2

## Features

* Completely unobtrusive:
  * Implemented as an [Application-Level Add-in](http://msdn.microsoft.com/bb386298.aspx), which can be used together with a [Document-Level Customization](http://msdn.microsoft.com/zcfbd2sk.aspx) and other Application-Level Add-Ins.
  * Allows to enable and disable each automatism separately.
* Allows to enforce copying of all styles from the attached Microsoft Word template into the active Microsoft Word document, overwriting any existing styles in the document that have the same name.
* TBD

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://gnu.org/licenses/lgpl.txt>.
