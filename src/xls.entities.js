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

var XlsDocument = function () {
	this.Author = "";
	this.LastAuthor = "";
	this.Created = "";
	this.Company = "";
	this.Version = "12.00";
};

var XlsWorkbook = function () {
	this.WindowHeight = "";
	this.WindowWidth = "";
	this.WindowTopX = "";
	this.WindowTopY = "";
	this.IsProtectStructure = false;
	this.IsProtectWindows = false;
};

var XlsWorksheet = function () {
	this.ID = "";
	this.Name = "";
	this.FullColumns = "";
	this.FullRows = "";
	this.IsProtectObjects = false;
	this.IsProtectScenarios = false;
	this.Columns = function (i) {
		this.Count = 0;
		this.Width = 0;
	};
	this.Rows = function (j) {
		this.Count = 0;
		this.Height = 0;
	};
	this.Cell = function (i, j) {
		this.Type = "";
		this.Value = "";
		this.Name = "";
	};
};

var XlsStyle = function (xlsFont) {
	this.Alignment = function () {
		this.Horizontal = XlsAlignmentType.Left;
		this.Vertical = XlsAlignmentType.Bottom;
	};
	this.Borders = "";
	this.Font = xlsFont;
	this.Interior = "";
	this.NumberFormat = "";
	this.Protection = "";
};

var XlsFont = function () {
	this.Name = "";
	this.Size = "";
};

var XlsName = function () {
	this.SheetName = "";
	this.CellName = "";
	this.CellAddress = "";
};

var XlsAlignmentType = {
	Left: "Left",
	Right: "Right",
	Center: "Center",
	Top: "Top",
	Bottom: "Bottom",
	Middle: "Middle"
};