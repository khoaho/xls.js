/**
 * xls.js JavaScript Library core
 * http://khoaho.github.com/xls.js/
 * Copyright (C) 2012 by Khoa Ho
 * Licensed under the MIT or GPL Version 2 licenses.
 * Date: Jun 03 2012
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

var XlsWriter = function () {
	"use strict";
	var xmlns = "urn:schemas-microsoft-com:office:spreadsheet",
		xmlns_o = "urn:schemas-microsoft-com:office:office",
		xmlns_x = "urn:schemas-microsoft-com:office:excel",
		xmlns_ss = "urn:schemas-microsoft-com:office:spreadsheet",
		xmlns_html = "http://www.w3.org/TR/REC-html40",
		_document,
		_documentProperties,
		_excelWorkbook,
		_styles,
		_names,
		_worksheets;

	function getDocumentProperties() {
		var element = _document.createElementNS(xmlns_o, "DocumentProperties"),
			attributes = getAttribute('Author', _documentProperties.Author)
				+ getAttribute('LastAuthor', _documentProperties.LastAuthor)
				+ getAttribute('Created', _documentProperties.Created)
				+ getAttribute('Company', _documentProperties.Company)
				+ getAttribute('Version', _documentProperties.Version);
		element.innerHTML = attributes;
		return element;
	}

	function getExcelWorkbook() {
		var element = _document.createElementNS(xmlns_x, "ExcelWorkbook"),
			attributes = getAttribute('WindowHeight', _excelWorkbook.WindowHeight)
				+ getAttribute('WindowWidth', _excelWorkbook.WindowWidth)
				+ getAttribute('WindowTopX', _excelWorkbook.WindowTopX)
				+ getAttribute('WindowTopY', _excelWorkbook.WindowTopY)
				+ getAttribute('ProtectStructure', _excelWorkbook.IsProtectStructure)
				+ getAttribute('ProtectWindows', _excelWorkbook.IsProtectWindows);
		element.innerHTML = attributes;
		return element;
	}

	function getStyles() {
		var length, i, element = _document.createElement('Styles');
		length = _styles.length;
		for (i = 0; i < length; i++) {
			element.appendChild(getStyle(_styles[i]));
		}
		return element;
	}

	function getStyle(style) {
		var element, alignment, borders, font, interior, numberFormat, protection;

		element = _document.createElement('Style');
		element.setAttribute('ss:ID', 'Default');
		element.setAttribute('ss:Name', 'Normal');

		alignment = _document.createElement('Alignment');
		alignment.setAttribute('ss:Vertical', style.Alignment.Vertical);

		borders = _document.createElement('Borders');
		borders.setAttribute('', style.Borders);

		font = _document.createElement('Font');
		font.setAttribute('ss:FontName', style.Font.Name);

		interior = _document.createElement('Interior');
		interior.setAttribute('', style.Interior);

		numberFormat = _document.createElement('NumberFormat');
		numberFormat.setAttribute('', style.NumberFormat);

		protection = _document.createElement('Protection');
		protection.setAttribute('', style.Protection);

		element.appendChild(alignment);
		element.appendChild(borders);
		element.appendChild(font);
		element.appendChild(interior);
		element.appendChild(numberFormat);
		element.appendChild(protection);
		return element;
	}

	function getNames() {
		var element = _document.createElement('Names'),
			key;
		for (key in _names) {
			element.appendChild(getName(_names[key]));
		}
		return element;
	}

	function getName(name) {
		var element, refersTo;
		element = _document.createElement('NamedRange');
		element.setAttribute('ss:Name', name.CellName);
		refersTo = "='" + name.SheetName + "'!" + name.CellAddress;
		element.setAttribute('ss:RefersTo', refersTo);
		return element;
	}

	function getWorksheets() {
		var worksheet, element, table, column, row, cellObj, cell, data,
			worksheetOptions, selected, protectObjects, protectScenarios,
			key, i, j;

		for (key in _worksheets) {
			worksheet = _worksheets[key];
			element = _document.createElement('Worksheet');
			element.setAttribute('ss:Name', worksheet.Name);

			table = _document.createElement('Table');
			table.setAttribute('ID', worksheet.ID);
			table.setAttribute('ss:ExpandedColumnCount', worksheet.Columns.Count);
			table.setAttribute('ss:ExpandedRowCount', worksheet.Rows.Count);
			table.setAttribute('ss:FullColumns', worksheet.FullColumns);
			table.setAttribute('ss:FullRows', worksheet.FullRows);

			for (i in worksheet.Columns.Count) {
				column = _document.createElement('Column');
				column.setAttribute('ss:Width', worksheet.Columns[i].Width);
				table.appendChild(column);
			}

			for (j in worksheet.Rows.Count) {
				row = _document.createElement('Row');
				for (i in worksheet.Columns.Count) {
					cellObj = worksheet.Cell[j, i];
					cell = _document.createElement('Cell');

					data = _document.createElement('Data');
					data.setAttribute('ss:Type', cellObj.Type);
					data.innerHTML = cellObj.Value;
					cell.appendChild(data);

					if (!!cellObj.Name) {
						var namedCell = _document.createElement('NamedCell');
						namedCell.setAttribute('ss:Name', cellObj.Name);
						cell.appendChild(namedCell);
					}

					row.appendChild(cell);
				}
				table.appendChild(row);
			}

			worksheet.appendChild(table);

			worksheetOptions = _document.createElementNS(xmlns_x, 'WorksheetOptions');
			selected = _document.createElement('Selected');
			protectObjects = _document.createElement('ProtectObjects');
			protectObjects.innerHTML = worksheet.IsProtectObjects;
			protectScenarios = _document.createElement('ProtectScenarios');
			protectScenarios.innerHTML = worksheet.IsProtectScenarios;

			worksheetOptions.appendChild(selected);
			worksheetOptions.appendChild(protectObjects);
			worksheetOptions.appendChild(protectScenarios);

			worksheet.appendChild(worksheetOptions);
		}
	}

	function getAttribute(name, value) {
		return "<" + name + ">" + value + "</" + name + ">";
	}

	function getWorkbook() {
		var workbook = _document.createElement('Workbook');
		workbook.setAttribute('xmlns', xmlns);
		workbook.setAttribute('xmlns:o', xmlns_o);
		workbook.setAttribute('xmlns:x', xmlns_x);
		workbook.setAttribute('xmlns:ss', xmlns_ss);
		workbook.setAttribute('xmlns:html', xmlns_html);

		_documentProperties = getDocumentProperties();
		_excelWorkbook = getExcelWorkbook();
		_styles = getStyles();
		_names = getNames();
		_worksheets = getWorksheets();

		workbook.appendChild(_documentProperties);
		workbook.appendChild(_excelWorkbook);
		workbook.appendChild(_styles);
		workbook.appendChild(_names);
		for (var key in _worksheets) {
			workbook.appendChild(_worksheets[key]);
		}
	}

	function writeFile() {
		return '<?xml version="1.0"?>' + '<?mso-application progid="Excel.Sheet"?>' + getWorkbook();
	}

	return {
		Save: function () {
			_document = document;
			writeFile();
		}
	};
}