/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

(function(){

    //------------------------------------------------------------------------------------------------------------------
	//
	// Internal
	//
	//------------------------------------------------------------------------------------------------------------------

    let FIELDS_HIGHLIGHT = {
        r: 204, 
        g: 215,
        b: 255
    }
    let BUTTON_PRESSED = {
        r: 153,
        g: 193,
        b: 218
    }
    
    let FIELD_TYPE = {
        button:         "button",
        checkbox:       "checkbox",
        combobox:       "combobox",
        listbox:        "listbox",
        radiobutton:    "radiobutton",
        signature:      "signature",
        text:           "text"
    }

    let ACTION_TRIGGER_TYPES = {
        MouseUp:    0,
        MouseDown:  1,
        MouseEnter: 2,
        MouseExit:  3,
        OnFocus:    4,
        OnBlur:     5,
        Keystroke:  6,
        Validate:   7,
        Calculate:  8,
        Format:     9
    }

    let BORDER_TYPES = {
        solid:      0,
        dashed:     1,
        beveled:    2,
        inset:      3,
        underline:  4
    }

    let LISTBOX_SELECTED_COLOR = {
        r: 153,
        g: 193,
        b: 218
    }

    //------------------------------------------------------------------------------------------------------------------
	//
	// pdf api types
	//
	//------------------------------------------------------------------------------------------------------------------
    
    let ALIGN_TYPE = {
        left:   "left",
        center: "center",
        right:  "right"
    }

    let border = {
        "s": "solid",
        "b": "beveled",
        "d": "dashed",
        "i": "inset",
        "u": "underline"
    }

    let position = {
        "textOnly":   0,
        "iconOnly":   1,
        "iconTextV":  2,
        "textIconV":  3,
        "iconTextH":  4,
        "textIconH":  5,
        "overlay":    6
    }

    let scaleHow = {
        "proportional":   0,
        "anamorphic":     1
    }

    let scaleWhen = {
        "always":     0,
        "never":      1,
        "tooBig":     2,
        "tooSmall":   3
    }

    const CHAR_LIM_MAX = 500; // to do проверить

    let display = {
        "visible":  0,
        "hidden":   1,
        "noPrint":  2,
        "noView":   3
    }

    // For Span attributes (start)
    let FONT_STRETCH = ["ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal",
        "semi-expanded", "expanded", "extra-expanded", "ultra-expanded"];

    let FONT_STYLE = {
        italic: "italic",
        normal: "normal"
    }

    let FONT_WEIGHT = [100, 200, 300, 400, 500, 600, 700, 800, 900];

    // for CSpan (end)

    
    // default availible colors
    let color = {
        "transparent":  [ "T" ],
        "black":        [ "G", 0 ],
        "white":        [ "G", 1 ],
        "red":          [ "RGB", 1,0,0 ],
        "green":        [ "RGB", 0,1,0 ],
        "blue":         [ "RGB", 0, 0, 1 ],
        "cyan":         [ "CMYK", 1,0,0,0 ],
        "magenta":      [ "CMYK", 0,1,0,0 ],
        "yellow":       [ "CMYK", 0,0,1,0 ],
        "dkGray":       [ "G", 0.25 ],  // version 4.0
        "gray":         [ "G", 0.5 ],   // version 4.0
        "ltGray":       [ "G", 0.75 ]   // version 4.0
    }

    // please use copy of this object
    let DEFAULT_SPAN = {
        "alignment":        ALIGN_TYPE.left,
        "fontFamily":       ["sans-serif"],
        "fontStretch":      "normal",
        "fontStyle":        "normal",
        "fontWeight":       400,
        "strikethrough":    false,
        "subscript":        false,
        "superscript":      false,
        "text":             "",
        "color":            color["black"],
        "textSize":         12.0,
        "underline":        false
    }

    // Defines how a button reacts when a user clicks it.
    // The four highlight modes supported are:
    let highlight = {
        "n": "none",
        "i": "invert",
        "p": "push",
        "o": "outline"
    }
    
    let LINE_WIDTH = {
        "none":   0,
        "thin":   1,
        "medium": 2,
        "thick":  3
    }

    let VALID_ROTATIONS = [0, 90, 180, 270];

    // Allows the user to set the glyph style of a check box or radio button.
    // The glyph style is the graphic used to indicate that the item has been selected.
    let style = {
        "ch": "check",
        "cr": "cross",
        "di": "diamond",
        "ci": "circle",
        "st": "star",
        "sq": "square"
    }

    const MAX_TEXT_SIZE = 32767;

    // freeze objects
    Object.freeze(FIELDS_HIGHLIGHT);
    Object.freeze(FIELD_TYPE);
    Object.freeze(ALIGN_TYPE);
    Object.freeze(border);
    Object.freeze(position);
    Object.freeze(scaleHow);
    Object.freeze(scaleWhen);
    Object.freeze(FONT_STRETCH);
    Object.freeze(FONT_STYLE);
    Object.freeze(FONT_WEIGHT);
    Object.freeze(color);
    Object.freeze(highlight);
    Object.freeze(VALID_ROTATIONS);
    Object.freeze(style);
    Object.freeze(ACTION_TRIGGER_TYPES);

    // base form class with attributes and method for all types of forms
	function CBaseField(sName, sType, nPage, aRect)
    {
        this.type = sType;
        
        this._kids          = [];
        this._borderStyle   = BORDER_TYPES.solid;
        this._delay         = false;
        this._display       = display["visible"];
        this._doc           = null;
        this._fillColor     = ["RGB",1,1,1];
        this._bgColor       = undefined;          // prop for old versions (fillColor)
        this._hidden        = false;             // This property has been superseded by the display property and its use is discouraged.
        this._lineWidth     = LINE_WIDTH.thin;  // In older versions of this specification, this property was borderWidth
        this._borderWidth   = undefined;       
        this._name          = sName;         // to do
        this._page          = nPage;        // integer | array
        this._print         = true;        // This property has been superseded by the display property and its use is discouraged.
        this._readonly      = false;
        this._rect          = aRect;         // scaled rect
        this._origRect      = [];           // orig rect as in file
        this._required      = false;       // for all except button
        this._rotation      = 0;
        this._strokeColor   = ["T"];     // In older versions of this specification, this property was borderColor. The use of borderColor is now discouraged,
                                        // although it is still valid for backward compatibility.
        this._borderColor   = undefined;
        this._submitName    = "";
        this._textColor     = ["RGB",0,0,0];
        this._fgColor       = undefined;
        this._textSize      = 10; // 0 == max text size // to do
        this._userName      = ""; // It is intended to be used as tooltip text whenever the cursor enters a field. 
        //It can also be used as a user-friendly name, instead of the field name, when generating error messages.

        this._actions = new CFormActions();

        // internal
        this._contentRect = {
            X: 0,
            Y: 0,
            W: 0,
            H: 0,
            Page: nPage
        }
        this._formRect = {
            X: 0,
            Y: 0,
            W: 0,
            H: 0,
            Page: nPage
        }

        this._oldContentPos = {X: 0, Y: 0, XLimit: 0, YLimit: 0};
        this._curShiftView = { // смещение, когда мы скролим, т.е. активное смещение
            x: 0,
            y: 0
        }
        this._originShiftView = { // смещение, когда значение формы применено (т.е. форма не активна)
            x: 0,
            y: 0
        }

        this._apIdx = -1; // индекс формы на странице в исходном файле (в массиве метода getInteractiveForms), используется для получения appearance
        
        this._needDrawHighlight = true;
        this._needRecalc        = false;
        this._wasChanged        = false; // была ли изменена форма
        this._bDrawFromStream   = false; // нужно ли рисовать из стрима

        private_getViewer().ImageMap = {};
        private_getViewer().InitDocument = function() {return};

        this._partialName = sName;
        this._apiForm = this.GetFormApi();
    }
    
    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.GetField = function(sName) {
        for (let i = 0; i < this._kids.length; i++) {
            if (this._kids[i]._partialName == sName)
                return this._kids[i];
        }

        return null;
    };

    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.AddKid = function(oField) {
        this._kids.push(oField);
        oField._parent = this;
    };
    /**
	 * Removes field from kids.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @param {CBaseField} oField - the field to remove.
	 * @returns {boolean} - returns false if field isn't in the field kids.
	 */
    CBaseField.prototype.RemoveKid = function(oField) {
        let nIndex = this._kids.indexOf(oField);
        if (nIndex != -1) {
            this._kids.splice(nIndex, 1);
            oField._parent = null;
            return true;
        }

        return false;
    };

    CBaseField.prototype.getFormRelRect = function()
    {
        return this._contentRect;
    };
    CBaseField.prototype.getFormRect = function()
    {
        return this._formRect;
    };
    
    CBaseField.prototype.IntersectWithRect = function(X, Y, W, H)
    {
        let oViewer     = private_getViewer();
        var arrRects    = [];
        var oBounds     = this.getFormRelRect();

        let nPageIndX = oViewer.pageDetector.pages[0].x / AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_pix_to_mm;
        let nPageIndY = oViewer.pageDetector.pages[0].y / AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_pix_to_mm;

        var nLeft   = Math.max(X, oBounds.X) + nPageIndX / oViewer.zoom;
        var nRight  = Math.min(X + W, oBounds.X + oBounds.W) + nPageIndX / oViewer.zoom;
        var nTop    = Math.max(Y, oBounds.Y) + nPageIndY / oViewer.zoom;
        var nBottom = Math.min(Y + H, oBounds.Y + oBounds.H) + nPageIndY / oViewer.zoom;

        if (nLeft < nRight && nTop < nBottom)
        {
            arrRects.push({
                X : nLeft,
                Y : nTop,
                W : nRight - nLeft,
                H : nBottom - nTop
            });
        }

        return arrRects;
    };
    CBaseField.prototype.GetFullName = function() {
        if (this._parent)
        {
            if (this._partialName != "")
                return `${this._parent.GetFullName()}.${this._partialName}`
            else
                return this._parent.GetFullName();
        }

        return this._partialName ? this._partialName : "";
    };
    /**
	 * Sets the JavaScript action of the field for a given trigger.
     * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof CBaseField
     * @param {cTrigger} cTrigger - A string that sets the trigger for the action.
     * @param {string} cScript - The JavaScript code to be executed when the trigger is activated.
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetAction = function(cTrigger, cScript) {
        switch (cTrigger) {
            case "MouseUp":
                this._actions.MouseUp = new CFormAction(ACTION_TRIGGER_TYPES.MouseUp, cScript);
                break;
            case "MouseDown":
                this._actions.MouseDown = new CFormAction(ACTION_TRIGGER_TYPES.MouseDown, cScript);
                break;
            case "MouseEnter":
                this._actions.MouseEnter = new CFormAction(ACTION_TRIGGER_TYPES.MouseEnter, cScript);
                break;
            case "MouseExit":
                this._actions.MouseExit = new CFormAction(ACTION_TRIGGER_TYPES.MouseExit, cScript);
                break;
            case "OnFocus":
                this._actions.OnFocus = new CFormAction(ACTION_TRIGGER_TYPES.OnFocus, cScript);
                break;
            case "OnBlur":
                this._actions.OnBlur = new CFormAction(ACTION_TRIGGER_TYPES.OnBlur, cScript);
                break;
            case "Keystroke":
                this._actions.Keystroke = new CFormAction(ACTION_TRIGGER_TYPES.Keystroke, cScript);
                break;
            case "Validate":
                this._actions.Validate = new CFormAction(ACTION_TRIGGER_TYPES.Validate, cScript);
                break;
            case "Calculate":
                this._actions.Calculate = new CFormAction(ACTION_TRIGGER_TYPES.Calculate, cScript);
                break;
            case "Format":
                this._actions.Format = new CFormAction(ACTION_TRIGGER_TYPES.Format, cScript);
                break;
        }
    };

    CBaseField.prototype.DrawHighlight = function(oCtx) {
        let oViewer     = private_getViewer();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;

        let X       = Math.round(this._pagePos.x * nScale);
        let Y       = Math.round(this._pagePos.y * nScale);
        let nWidth  = Math.round(this._pagePos.w * nScale);
        let nHeight = Math.round(this._pagePos.h * nScale);
        
        oCtx.globalAlpha = 0.6; 
        if (this.type == "button" && this._buttonPressed) {
            oCtx.fillStyle = `rgb(${BUTTON_PRESSED.r}, ${BUTTON_PRESSED.g}, ${BUTTON_PRESSED.b})`;
            oCtx.fillRect(X, Y, nWidth, nHeight);
        }
        else {
            oCtx.fillStyle = `rgb(${FIELDS_HIGHLIGHT.r}, ${FIELDS_HIGHLIGHT.g}, ${FIELDS_HIGHLIGHT.b})`;
            oCtx.fillRect(X, Y, nWidth, nHeight);
        }
    };

    CBaseField.prototype.DrawBorders = function(oCtx) {
        let oViewer = private_getViewer();
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        let nLineWidth = 1 * nScale * this._lineWidth;
        oCtx.lineWidth = nLineWidth;
        if (nLineWidth == 0) {
            return;
        }

        let X       = Math.round(this._pagePos.x * nScale) + nLineWidth / 2;
        let Y       = Math.round(this._pagePos.y * nScale) + nLineWidth / 2;
        let nWidth  = Math.round(this._pagePos.w * nScale) - nLineWidth;
        let nHeight = Math.round(this._pagePos.h * nScale) - nLineWidth;

        let color = {
            r: this._borderColor[0] * 255,
            g: this._borderColor[1] * 255,
            b: this._borderColor[2] * 255
        }
        oCtx.strokeStyle = `rgb(${color.r}, ${color.g}, ${color.b})`;

        switch (this._borderStyle) {
            case BORDER_TYPES.solid:
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                break;
            case BORDER_TYPES.beveled:
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                oCtx.closePath();

                // bottom part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                oCtx.fillStyle = "rgb(192, 192, 192)";
                oCtx.closePath();
                oCtx.fill();

                // right part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                oCtx.fillStyle = "rgb(192, 192, 192)";
                oCtx.closePath();
                oCtx.fill();

                break;
            case BORDER_TYPES.dashed:
                oCtx.setLineDash([5 * oViewer.zoom]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                break;
            case BORDER_TYPES.inset:
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                oCtx.closePath();

                // left part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                // top part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                // bottom part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                oCtx.fillStyle = "rgb(191, 191, 191)";
                oCtx.closePath();
                oCtx.fill();

                // right part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                oCtx.fillStyle = "rgb(191, 191, 191)";
                oCtx.closePath();
                oCtx.fill();

                break;
            case BORDER_TYPES.underline:
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.moveTo(X, Y + nHeight);
                oCtx.lineTo(X + nWidth, Y + nHeight);
                oCtx.stroke();
                break;
        }

        // draw comb cells
        if ((this._borderStyle == BORDER_TYPES.solid || this._borderStyle == BORDER_TYPES.dashed) && (this.type == "text" && this._comb == true)) {
            let nCombWidth = (nWidth / this._charLimit);
            let nIndentX = nCombWidth;
            
            for (let i = 0; i < this._charLimit - 1; i++) {
                oCtx.moveTo(X + nIndentX, Y);
                oCtx.lineTo(X + nIndentX, Y + nHeight);
                oCtx.stroke();
                nIndentX += nCombWidth;
            }
        }
    };
    CBaseField.prototype.SetNeedRecalc = function(bRecalc) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            this._needRecalc = true;
            this.AddToRedraw();
        }
    };
    CBaseField.prototype.SetWasChanged = function(isChanged) {
        this._wasChanged = isChanged;
        this.SetDrawFromStream(!isChanged);
    };
    CBaseField.prototype.IsChanged = function() {
        return this._wasChanged;  
    };
    CBaseField.prototype.IsNeedDrawFromStream = function() {
        return this._bDrawFromStream;
    };
    CBaseField.prototype.SetDrawFromStream = function(bFromStream) {
        this._bDrawFromStream = bFromStream;
    };

    CBaseField.prototype.AddToRedraw = function() {
        let oViewer = private_getViewer();
        oViewer.pagesInfo.pages[this._page].needRedrawForms = true;
    };

    CBaseField.prototype.GetType = function() {
        return this.type;
    };

    CBaseField.prototype.SetReadOnly = function(bReadOnly) {
        this._readonly = bReadOnly;
    };
    
    CBaseField.prototype.SetRequired = function(bRequired) {
        if (this.type != "button")
            this._required = bRequired;
    };
    CBaseField.prototype.SetBorderColor = function(aColor) {
        this._borderColor = aColor;
    };


    /**
	 * Gets Api class for this form.
	 * @memberof CTextField
     * @param {number} nIdx - The 0-based index of the item in the list or -1 for the last item in the list.
     * @param {boolean} [bExportValue=true] - Specifies whether to return an export value.
	 * @typeofeditors ["PDF"]
     * @returns {ApiBaseField}
	 */
    CBaseField.prototype.GetFormApi = function() {
        if (this._apiForm)
            return this._apiForm;

        switch (this.type) {
            case "text":
                return new AscPDFEditor.ApiTextField(this);
            case "combobox":
                return new AscPDFEditor.ApiComboBoxField(this);
            case "listbox":
                return new AscPDFEditor.ApiListBoxField(this);
            case "checkbox":
                return new AscPDFEditor.ApiCheckBoxField(this);
            case "radiobutton":
                return new AscPDFEditor.ApiRadioButtonField(this);
            case "button":
                return new AscPDFEditor.ApiPushButtonField(this);
        }
    };

          
    function CPushButtonField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.button, nPage, aRect);

        this._buttonAlignX      = 50; // must be integer
        this._buttonAlignY      = 50; // must be integer
        this._buttonFitBounds   = undefined;
        this._buttonPosition    = position["textOnly"];
        this._buttonScaleHow    = undefined;
        this._highlight         = highlight["p"];
        this._textFont          = "ArialMT";

        this._buttonCaption     = undefined;
        this._downCaption       = undefined;
        this._rollOverCaption   = undefined;

        this._buttonPressed = false;
        // internal
        TurnOffHistory();
        this._content           = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);
        this.SetAlign(align_Center);

        this._captionRun            = null;
        this._downCaptionRun        = null;
        this._rollOverCaptionRun    = null; 
    }
    CPushButtonField.prototype = Object.create(CBaseField.prototype);
	CPushButtonField.prototype.constructor = CPushButtonField;
    CPushButtonField.prototype.AddImage = function(oImgData) {
		if (!oImgData) {
			return;
		}
		const oHTMLImg = oImgData.Image; m
		if (!oHTMLImg || oHTMLImg.width === 0 || oHTMLImg.height === 0 || this._buttonPosition == position["textOnly"]) {
			return;
		}

        let oExistDrawing = this.GetDrawing();
        if (oExistDrawing) {
            oExistDrawing.PreDelete();
            var oParentRun = oExistDrawing.GetRun();
            oParentRun.RemoveElement(oExistDrawing);
        }
        
        let oViewer = private_getViewer();
        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;
        let oMargins = this.GetBordersWidth();

		const dImgW = Math.max((oHTMLImg.width * AscCommon.g_dKoef_pix_to_mm), 1);
		const dImgH = Math.max((oHTMLImg.height * AscCommon.g_dKoef_pix_to_mm), 1);
		const oRect = this.getFormRelRect();
        let nContentWidth;
        switch (this._buttonPosition) {
            case position["iconTextH"]:
                nContentWidth = this._content.GetElement(0).GetContentWidthInRange();
                break;
            case position["textIconH"]:
                nContentWidth = this._content.GetElement(0).GetContentWidthInRange();
                break;
            default:
                nContentWidth = 0;
                break;
        }
		const dFrmW = oRect.W - ((oMargins.left + oMargins.right) * g_dKoef_pix_to_mm / scaleCoef + nContentWidth);
		const dFrmH = oRect.H - ((oMargins.top + oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef + nContentWidth);
		const dCW = dFrmW/dImgW;
		const dCH = dFrmH/dImgH;
		const dCoef = Math.min(dCW, dCH);
		const dDrawingW = dCoef*dImgW;
		const dDrawingH = dCoef*dImgH;
		const oDrawing = new AscCommonWord.ParaDrawing(dDrawingW, dDrawingH, null, this._content.DrawingDocument, this._content, null);
		
        let oShapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, this._content.Get_Theme(), null, null, null, 0);
		oShapeTrack.track({}, dDrawingW, dDrawingH);
		let oShape = oShapeTrack.getShape(true, this._content.DrawingDocument, null);
		oShape.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oShape);
        var oBodyPr = new AscFormat.CBodyPr();
        oBodyPr.setAnchor(1);
        oShape.setBodyPr(oBodyPr);
		oShape.createTextBoxContent();
        oShape.textBoxContent.SetParagraphAlign(AscCommon.align_Center);
        oShape.textBoxContent.GetElement(0).private_CompileParaPr(true);
        oShape.recalculateTxBoxContent();
        
        let oFill = new AscFormat.CUniFill();
        oFill.fill = new AscFormat.CBlipFill();
        oFill.fill.setRasterImageId(oImgData.src);
        oFill.fill.tile = null;
        oFill.fill.srcRect = null;
        oFill.fill.stretch = true;
        oFill.convertToPPTXMods();
        oShape.setFill(oFill);

		//oShape.spPr.setFill(oFill.UniFill);
		oShape.spPr.setLn(new AscFormat.CreateNoFillLine());

        let oRunForImg;
        let nContentH = this._content.Get_EmptyHeight();
        let oTargetPara;
        switch (this._buttonPosition) {
            case position["iconOnly"]:
                oRunForImg = this._content.GetElement(0).GetElement(0);
                break;
            case position["iconTextV"]:
                oRunForImg = this._content.GetElement(0).GetElement(0);
                break;
            case position["textIconV"]:
                oRunForImg = this._content.GetElement(1).GetElement(0);
                break;
            case position["iconTextH"]:
                oTargetPara = this._content.GetElement(0);
                if (oTargetPara.GetElementsCount() == 1) {
                    let oRun = new ParaRun(oTargetPara, false);
                    oTargetPara.Add_ToContent(oTargetPara.Content.length - 1, oRun);
                }

                oRunForImg = oTargetPara.GetElement(0);
                oRunForImg.Pr.Position = -(dDrawingH / 2 - nContentH / 4);
                oRunForImg.RecalcInfo.TextPr = true;
                oRunForImg.Get_CompiledPr();
                break;
            case position["textIconH"]:
                oTargetPara = this._content.GetElement(0);
                if (oTargetPara.GetElementsCount() == 1) {
                    let oRun = new ParaRun(oTargetPara, false);
                    oTargetPara.Add_ToContent(oTargetPara.Content.length - 1, oRun);
                }

                oRunForImg = oTargetPara.GetElement(1);
                oRunForImg.Pr.Position = -(dDrawingH / 2 - nContentH / 4);
                oRunForImg.RecalcInfo.TextPr = true;
                oRunForImg.Get_CompiledPr();
                break;
            case position["overlay"]:
                oTargetPara = this._content.GetElement(0);
                oTargetPara.Remove_FromContent(0, oTargetPara.GetElementsCount(), true);
                oTargetPara.CorrectContent();
                oRunForImg = oTargetPara.GetElement(0);

                if (this._buttonCaption) {
                    let oShapeCont = oShape.getDocContent();
                    let oCaptionRun = oShapeCont.GetElement(0).GetElement(0);
                    oCaptionRun.AddText(this._buttonCaption);
                }
                break;
        }

        oRunForImg.Add_ToContent(oRunForImg.Content.length, oDrawing);
		oDrawing.Set_Parent(oRunForImg);

        oShape.recalculate();
		this.SetNeedRecalc(true);
	};
    /**
	 * Gets the caption associated with a button.
	 * @memberof CPushButtonField
     * @param {number} [nFace=0] - (optional) If specified, gets a caption of the given type:
     * 0 — (default) normal caption
     * 1 — down caption
     * 2 — rollover caption
	 * @typeofeditors ["PDF"]
	 */
    CPushButtonField.prototype.GetCaption = function(nFace) {
        if (nFace == null)
            nFace = 0;

        switch (nFace) {
            case 0:
                return this._buttonCaption;
            case 1:
                return this._downCaption;
            case 2:
                return this._rollOverCaption;
        }

        return undefined;
    };
    /**
	 * Gets the caption associated with a button.
	 * @memberof CPushButtonField
     * @param {string} cCaption - The caption associated with the button.
     * @param {number} [nFace=0] - (optional) If specified, gets a caption of the given type:
     * 0 — (default) normal caption
     * 1 — down caption
     * 2 — rollover caption
	 * @typeofeditors ["PDF"]
	 */
    CPushButtonField.prototype.SetCaption = function(cCaption, nFace) {
        if (nFace == null)
            nFace = 0;

        if (cCaption == "" || typeof(cCaption) != "string")
            return false;
            
        switch (nFace) {
            case 0:
                this._buttonCaption = cCaption;
                let oCaptionRun;
                switch (this._buttonPosition) {
                    case position["textIconV"]:
                    case position["textOnly"]:
                    case position["textIconH"]:
                        oCaptionRun = this._content.GetElement(0).GetElement(0);
                        oCaptionRun.ClearContent();
                        oCaptionRun.AddText(cCaption);
                        this._captionRun = oCaptionRun;
                        break;
                    case position["iconOnly"]:
                        this._captionRun = null;
                        break;
                    case position["iconTextV"]:
                        oCaptionRun = this._content.GetElement(1).GetElement(0);
                        oCaptionRun.ClearContent();
                        oCaptionRun.AddText(cCaption);
                        this._captionRun = oCaptionRun;
                        break;
                    case position["iconTextH"]:
                        oCaptionRun = this._content.GetElement(0).GetElement(1);
                        oCaptionRun.ClearContent();
                        oCaptionRun.AddText(cCaption);
                        this._captionRun = oCaptionRun;
                        break;
                    case position["overlay"]:
                        let oDrawing = this.GetDrawing();
                        if (oDrawing) {
                            let oShapeCont = oDrawing.GraphicObj.getDocContent();
                            oCaptionRun = oShapeCont.GetElement(0).GetElement(0);
                        }
                        else {
                            oCaptionRun = this._content.GetElement(0).GetElement(0);
                        }
                        
                        oCaptionRun.ClearContent();
                        oCaptionRun.AddText(cCaption);
                        this._captionRun = oCaptionRun;
                        break;
                }
                break;
            case 1:
                this._downCaption = cCaption;
                break;
            case 2:
                this._rollOverCaption = cCaption;
                break;
        }
    };
    CPushButtonField.prototype.Draw = function(oCtx) {
        let oViewer = private_getViewer();

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = (this._rect[2] - this._rect[0]);
        let nHeight = (this._rect[3] - this._rect[1]);

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        this.DrawBorders(oCtx);
        let oMargins = this.GetBordersWidth();

        let contentX = (X + nWidth * 0.01 + oMargins.left) * g_dKoef_pix_to_mm;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit = (X + nWidth * 0.99 - oMargins.right) * g_dKoef_pix_to_mm;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        contentY = Y * g_dKoef_pix_to_mm;

        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        this._contentRect.X = contentX;
        this._contentRect.Y = contentY;
        this._contentRect.W = contentXLimit - contentX;
        this._contentRect.H = contentYLimit - contentY;

        // контент посередине
        let oDrawing = this.GetDrawing();
        let nContentH = this._content.Get_EmptyHeight();
        if (oDrawing) {
            if (this._content.Content.length == 2)
                contentY = contentY + ((contentYLimit - contentY) - (oDrawing.Extent.H + nContentH)) / 2;
            else
                contentY = contentY + ((contentYLimit - contentY) - oDrawing.Extent.H) / 2;
        }
        else {
            if (this._content.Content.length == 2) {
                if (this._buttonPosition == position["iconTextV"])
                    contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2 - nContentH;
                else
                    contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2;
            }
            else
                contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2;
        }
            

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
            if (oDrawing && this._buttonPosition == position["overlay"]) {
                oDrawing.GraphicObj.recalculateTxBoxContent();
            }
        }
        else if (this._needRecalc) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            if (oDrawing && this._buttonPosition == position["overlay"]) {
                oDrawing.GraphicObj.recalculateTxBoxContent();
            }

            this.SetNeedRecalc(false);
        }

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        oGraphics.init(oCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && (oViewer.fieldFillingMode || this.type == "combobox"))
            this._content.RecalculateCurPos();
        
        oGraphics.RemoveClip();
    };
    CPushButtonField.prototype.onMouseDown = function() {
        let oViewer = private_getViewer();
        this._buttonPressed = true;
        this._needDrawHighlight = true;

        oViewer._paintFormsHighlight();
    };
    CPushButtonField.prototype.onMouseUp = function() {
        let oViewer = private_getViewer();
        this._buttonPressed = false;
        this._needDrawHighlight = true;

        this.buttonImportIcon();
        oViewer._paintFormsHighlight();
    };
    CPushButtonField.prototype.buttonImportIcon = function() {
        let Api = editor;
        let oThis = this;

        Api.oSaveObjectForAddImage = this;
        AscCommon.ShowImageFileDialog(Api.documentId, Api.documentUserId, undefined, function(error, files)
		{
			Api._uploadCallback(error, files, oThis);
		}, function(error)
		{
			if (c_oAscError.ID.No !== error)
			{
				Api.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
			}
			if (obj && obj.sendUrlsToFrameEditor && Api.isOpenedChartFrame)
			{
				Api.sendStartUploadImageActionToFrameEditor();
			}
			Api.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
		});
    };
    CPushButtonField.prototype.GetDrawing = function() {
        return this._content.GetAllDrawingObjects()[0];
    };
    /**
	 * Controls how the text and the icon of the button are positioned with respect to each other within the button face. The
     * convenience position object defines all of the valid alternatives..
	 * @memberof CPushButtonField
     * @param {number} nType
	 * @typeofeditors ["PDF"]
	 */
    CPushButtonField.prototype.SetLayout = function(nType) {
        switch (nType) {
            case position["textOnly"]:
                this.SetTextOnly();
                break;
            case position["iconOnly"]:
                this.SetIconOnly();
                break;
            case position["iconTextV"]:
                this.SetIconTextV();
                break;
            case position["textIconV"]:
                this.SetTextIconV();
                break;
            case position["iconTextH"]:
                this.SetIconTextH();
                break;
            case position["textIconH"]:
                this.SetTextIconH();
                break;
            case position["overlay"]:
                this.SetOverlay();
                break;
        }
    };
    CPushButtonField.prototype.SetTextOnly = function() {
        if (this._buttonPosition == position["textOnly"])
            return;

        this._buttonPosition = position["textOnly"];

        let oPara;
        if (this._content.Content.length == 2) {
            for (let i = 0; i < this._content.Content.length; i++) {
                oPara = this._content.GetElement(i);
    
                if (oPara.GetAllDrawingObjects().length > 0) {
                    this._content.RemoveFromContent(i, 1, false);
                    break;
                }
            }
        }
        else {
            if (this._buttonCaption) {
                oPara = this._content.GetElement(0);
                oPara.ClearContent();
                oPara.CheckParaEnd();
                oPara.GetElement(0).AddText(this._buttonCaption);
                this._captionRun = oPara.GetElement(0);
            }
        }
    };
    CPushButtonField.prototype.SetIconOnly = function() {
        if (this._buttonPosition == position["iconOnly"])
            return;

        this._buttonPosition = position["iconOnly"];

        let oPara;
        if (this._content.Content.length == 2) {
            for (let i = 0; i < this._content.Content.length; i++) {
                oPara = this._content.GetElement(i);
    
                if (oPara.GetAllDrawingObjects().length == 0) {
                    this._content.RemoveFromContent(i, 1, false);
                    break;
                }
            }
        }
        else {
            oPara = this._content.GetElement(0);
            
            let aRunForDel = [];
            let oTmpRun;
            for (let i = 0; i < oPara.Content.length - 1; i++) {
                oTmpRun = oPara.GetElement(i);
                if (oTmpRun.GetAllDrawingObjects().length == 0 && oTmpRun.Content.length != 0) {
                    aRunForDel.push(oTmpRun);
                }
            }

            let nPosInPara;
            for (let oRun of aRunForDel) {
                oRun.PreDelete();
                nPosInPara = oPara.Content.indexOf(oRun);
                oPara.Remove_FromContent(nPosInPara, 1, true);
            }
            oPara.CorrectContent();
        }
    };
    CPushButtonField.prototype.SetIconTextV = function() {
        if (this._buttonPosition == position["iconTextV"])
            return;

        let oPara1;
        let oPara2;
        if (this._content.Content.length == 2) {
            oPara1 = this._content.GetElement(0);
            oPara2 = this._content.GetElement(1);

            if (oPara2.GetAllDrawingObjects().length != 0) {
                [this._content.Content[0], this._content.Content[1]] = [this._content.Content[1], this._content.Content[0]];
                oPara1.Set_DocumentIndex(1);
                oPara2.Set_DocumentIndex(0);
            }
        }
        else {
            oPara1 = this._content.GetElement(0);
            let aRunForDel = [];
            let oTmpRun;
            for (let i = 0; i < oPara1.Content.length - 1; i++) {
                oTmpRun = oPara1.GetElement(i);
                if (oTmpRun.GetAllDrawingObjects().length == 0 && oTmpRun.Content.length != 0) {
                    aRunForDel.push(oTmpRun);
                }
            }

            let nPosInPara;
            for (let oRun of aRunForDel) {
                oRun.PreDelete();
                nPosInPara = oPara1.Content.indexOf(oRun);
                oPara1.Remove_FromContent(nPosInPara, 1, true);
            }

            let oNewPara = new AscCommonWord.Paragraph(oPara1.DrawingDocument, this._content, false);
            oNewPara.Correct_Content();
            this._content.AddToContent(1, oNewPara);
            oNewPara.Set_Align(align_Center);
            oNewPara.private_CompileParaPr(true);
        }

        this._buttonPosition = position["iconTextV"];
    };
    CPushButtonField.prototype.SetTextIconV = function() {
        if (this._buttonPosition == position["textIconV"])
            return;

        let oPara1;
        let oPara2;
        if (this._content.Content.length == 2) {
            oPara1 = this._content.GetElement(0);
            oPara2 = this._content.GetElement(1);

            if (oPara1.GetAllDrawingObjects().length != 0) {
                [this._content.Content[0], this._content.Content[1]] = [this._content.Content[1], this._content.Content[0]];
                oPara1.Set_DocumentIndex(1);
                oPara2.Set_DocumentIndex(0);
            }
        }
        else {
            oPara1 = this._content.GetElement(0);
            let aRunForDel = [];
            let oTmpRun;
            for (let i = 0; i < oPara1.Content.length - 1; i++) {
                oTmpRun = oPara1.GetElement(i);
                if (oTmpRun.GetAllDrawingObjects().length == 0 && oTmpRun.Content.length != 0) {
                    aRunForDel.push(oTmpRun);
                }
            }

            let nPosInPara;
            for (let oRun of aRunForDel) {
                oRun.PreDelete();
                nPosInPara = oPara1.Content.indexOf(oRun);
                oPara1.Remove_FromContent(nPosInPara, 1, true);
            }

            let oNewPara = new AscCommonWord.Paragraph(oPara1.DrawingDocument, this._content, false);
            oNewPara.Correct_Content();
            this._content.AddToContent(0, oNewPara);
            oNewPara.Set_Align(align_Center);
            oNewPara.private_CompileParaPr(true);
        }

        this._buttonPosition = position["textIconV"];
    };
    CPushButtonField.prototype.SetIconTextH = function() {
        if (this._buttonPosition == position["iconTextH"])
            return;

        let oPara;
        if (this._content.Content.length == 2) {
            for (let i = 0; i < this._content.Content.length; i++) {
                oPara = this._content.GetElement(i);
                if (oPara.GetAllDrawingObjects().length == 0) {
                    this._content.RemoveFromContent(i, 1, false);
                    break;
                }
            }
        }

        oPara = this._content.GetElement(0);
        let aRunForDel = [];
        let oTmpRun;
        
        for (let i = 0; i < oPara.Content.length - 1; i++) {
            oTmpRun = oPara.GetElement(i);
            if (oTmpRun.GetAllDrawingObjects().length == 0 && oTmpRun.Content.length != 0) {
                aRunForDel.push(oTmpRun);
            }
        }

        let nPosInPara;
        for (let oRun of aRunForDel) {
            oRun.PreDelete();
            nPosInPara = oPara.Content.indexOf(oRun);
            oPara.Remove_FromContent(nPosInPara, 1, true);
        }
        oPara.CorrectContent();

        if (this._buttonCaption) {
            oTmpRun = new ParaRun(oPara, false);
            oTmpRun.AddText(this._buttonCaption);
            oPara.Add_ToContent(oPara.Content.length - 1, oTmpRun);
        }

        this._buttonPosition = position["iconTextH"];
    };
    CPushButtonField.prototype.SetTextIconH = function() {
        if (this._buttonPosition == position["textIconH"])
            return;

        let oPara;
        if (this._content.Content.length == 2) {
            for (let i = 0; i < this._content.Content.length; i++) {
                oPara = this._content.GetElement(i);
                if (oPara.GetAllDrawingObjects().length == 0) {
                    this._content.RemoveFromContent(i, 1, false);
                    break;
                }
            }
        }

        oPara = this._content.GetElement(0);
        let aRunForDel = [];
        let oTmpRun;
        
        for (let i = 0; i < oPara.Content.length - 1; i++) {
            oTmpRun = oPara.GetElement(i);
            if (oTmpRun.GetAllDrawingObjects().length == 0 && oTmpRun.Content.length != 0) {
                aRunForDel.push(oTmpRun);
            }
        }

        let nPosInPara;
        for (let oRun of aRunForDel) {
            oRun.PreDelete();
            nPosInPara = oPara.Content.indexOf(oRun);
            oPara.Remove_FromContent(nPosInPara, 1, true);
        }
        oPara.CorrectContent();

        if (this._buttonCaption) {
            oTmpRun = new ParaRun(oPara, false);
            oTmpRun.AddText(this._buttonCaption);
            oPara.Add_ToContent(0, oTmpRun);
        }
        
        this._buttonPosition = position["textIconH"];
    };
    CPushButtonField.prototype.SetOverlay = function() {
        if (this._buttonPosition == position["overlay"])
            return;

        let oPara;
        if (this._content.Content.length == 2) {
            for (let i = 0; i < this._content.Content.length; i++) {
                oPara = this._content.GetElement(i);
                if (oPara.GetAllDrawingObjects().length == 0) {
                    this._content.RemoveFromContent(i, 1, false);
                    break;
                }
            }
        }

        oPara = this._content.GetElement(0);
        let aRunForDel = [];
        let oTmpRun;
        
        for (let i = 0; i < oPara.Content.length - 1; i++) {
            oTmpRun = oPara.GetElement(i);
            if (oTmpRun.GetAllDrawingObjects().length == 0) {
                aRunForDel.push(oTmpRun);
            }
        }

        let nPosInPara;
        for (let oRun of aRunForDel) {
            oRun.PreDelete();
            nPosInPara = oPara.Content.indexOf(oRun);
            oPara.Remove_FromContent(nPosInPara, 1, true);
        }
        oPara.CorrectContent();

        let oDrawing = this.GetDrawing();
        if (oDrawing && this._buttonCaption) {
            let oShapeCont = oDrawing.GraphicObj.getDocContent();
            oTmpRun = oShapeCont.GetElement(0).GetElement(0);
            oTmpRun.AddText(this._buttonCaption);
        }

        this._buttonPosition = position["overlay"];
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CPushButtonField
	 * @typeofeditors ["PDF"]
	 */
    CPushButtonField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._buttonAlignX      = aFields[i]._buttonAlignX;
                this._buttonAlignY      = aFields[i]._buttonAlignY;
                this._buttonFitBounds   = aFields[i]._buttonFitBounds;
                this._buttonPosition    = Object.assign(this._buttonPosition, aFields[i]._buttonPosition);
                this._buttonScaleHow    = aFields[i]._buttonScaleHow;
                this._highlight         = aFields[i]._highlight;
                this._textFont          = aFields[i]._textFont;

                this._actions = aFields[i]._actions ? aFields[i]._actions.Copy(this) : null;

                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();

                // format content
                oPara = this._contentFormat.GetElement(0);
                oParaToCopy = aFields[i]._contentFormat.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                break;
            }
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CPushButtonField
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CPushButtonField.prototype.ApplyValueForAll = function(bUnionPoints) {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThisPara = this._content.GetElement(0);
        
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        TurnOffHistory();

        if (bUnionPoints)
            this.UnionLastHistoryPoints();

        if (aFields.length == 1)
            this._needApplyToAll = false;

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i].SetNeedRecalc(true);
        }
    };

    function CBaseCheckBoxField(sName, sType, nPage, aRect)
    {
        CBaseField.call(this, sName, sType, nPage, aRect);

        this._value         = "Off";
        this._exportValue   = "Yes";
        
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        
        let oPara = this._content.GetElement(0);
        oPara.Recalculate_Page(0);
        oPara.SetParagraphAlign(align_Center);
        oPara.CompiledPr.NeedRecalc = false;
        oPara.CompiledPr.Pr.ParaPr.Jc = align_Center;
    }
    
    CBaseCheckBoxField.prototype = Object.create(CBaseField.prototype);
	CBaseCheckBoxField.prototype.constructor = CBaseCheckBoxField;

    CBaseCheckBoxField.prototype.Draw = function(oCtx) {
        let oViewer = private_getViewer();

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = (this._rect[2] - this._rect[0]);
        let nHeight = (this._rect[3] - this._rect[1]);

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        if (this.type == "checkbox")
            this.DrawBorders(oCtx);

        let oMargins = this.GetBordersWidth();

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        contentY = Y * g_dKoef_pix_to_mm;

        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        this._contentRect.X = contentX;
        this._contentRect.Y = contentY;
        this._contentRect.W = contentXLimit - contentX;
        this._contentRect.H = contentYLimit - contentY;

        // подгоняем размер галочки
        this.ProcessAutoFitContent();

        // выставляем текст посередине
        // let nContentH = this._content.GetElement(0).Get_EmptyHeight();
        // contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (this._needRecalc) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.SetNeedRecalc(false);
        }

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        oGraphics.init(oCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && (oViewer.fieldFillingMode || this.type == "combobox"))
            this._content.RecalculateCurPos();
        
        oGraphics.RemoveClip();
    };

    CBaseCheckBoxField.prototype.ProcessAutoFitContent = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        let oTextPr = oRun.Get_CompiledPr(true);
        let oBounds = this.getFormRelRect();

        g_oTextMeasurer.SetTextPr(oTextPr, null);
	    g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        var nTextHeight = g_oTextMeasurer.GetHeight();
	    var nMaxWidth   = oPara.RecalculateMinMaxContentWidth(false).Max;
	    var nFontSize   = oTextPr.FontSize;

        if (nMaxWidth < 0.001 || nTextHeight < 0.001 || oBounds.W < 0.001 || oBounds.H < 0.001)
		    return;

	    var nNewFontSize = nFontSize;

        nNewFontSize = oBounds.H / g_dKoef_pt_to_mm;
        oRun.SetFontSize(nNewFontSize);
    };
    CBaseCheckBoxField.prototype.SetExportValue = function(sValue) {
        this._exportValue = sValue;
    };
    
    // for radiobutton
    const CheckedSymbol   = 0x25C9;
	const UncheckedSymbol = 0x25CB;

    function CCheckBoxField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.checkbox, nPage, aRect);

        this._style = style.ch;
    }
    CCheckBoxField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CCheckBoxField.prototype.constructor = CCheckBoxField;

    CCheckBoxField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThis = this;
        aFields.forEach(function(field) {
            if (field == oThis) {
                CreateNewHistoryPointForField(oThis);
            }
            else
                TurnOffHistory();

            let oRun = field._content.GetElement(0).GetElement(0);
            if (field._value != "Off") {
                oRun.ClearContent();
                field._value = "Off";
            }
            else {
                field._value = field._exportValue;
                oRun.AddText('✓');
            }

            field.SetNeedRecalc(true);
        });
        
        private_getViewer()._paintForms();
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.ApplyValueForAll = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThisPara = this._content.GetElement(0);
        let oThisRun = oThisPara.GetElement(0);

        this.CheckValue();
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
             // пропускаем текущее поле, т.к. уже были изменения после redo/undo
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oFieldRun = oFieldPara.GetElement(0);

            oFieldRun.ClearContent();
            for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
            }

            aFields[i].SetNeedRecalc(true);
            aFields[i].CheckValue();
        }
    };
    /**
	 * Checks value of the field and corrects it.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.CheckValue = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        if (oRun.GetText() != "")
            this._value = this._exportValue;
        else
            this._value = "Off";
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let nThisIdx = aFields.indexOf(this);
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                if (aFields[i]._value != "Off")
                    this._value = this._exportValue;
                break;
            }
        }
    };

    CCheckBoxField.prototype.SetValue = function(sValue) {
        if (sValue != "Off") {
            let oRun = this._content.GetElement(0).GetElement(0);
            oRun.AddText('✓');
            this._value = sValue;
        }
    };

    function CRadioButtonField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.radiobutton, nPage, aRect);
        
        let oRun = this._content.GetElement(0).GetElement(0);
        //oRun.AddText(String.fromCharCode(UncheckedSymbol));
        oRun.AddText("〇");

        this._radiosInUnison = false;
        this._noToggleToOff = true;

        this._style = style.ci;
        this._needDrawHighlight = false;
    }
    CRadioButtonField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
                
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._radiosInUnison = aFields[i]._radiosInUnison;

                if (aFields[i]._value != "Off")
                    this._value = this._exportValue;

                if (this._radiosInUnison && this._exportValue == aFields[i]._exportValue) {
                    let oPara = this._content.GetElement(0);
                    let oParaToCopy = aFields[i]._content.GetElement(0);

                    oPara.ClearContent();
                    for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                        oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                    }
                    oPara.CheckParaEnd();
                
                    break;
                }
            }
        }
    };
    CRadioButtonField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThis = this;

        CreateNewHistoryPointForField(oThis);
        if (false == this._radiosInUnison) {
            if (this._value != "Off") {
                if (this._noToggleToOff == false) {
                    this.SetChecked(false);
                    this.SetNeedRecalc(true);
                }
                if (AscCommon.History.Is_LastPointEmpty())
                    AscCommon.History.Remove_LastPoint();

                return;
            }
            else {
                this.SetChecked(true);
                this.SetNeedRecalc(true);
            }

            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field._value != "Off") {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
            }); 
        }
        else {
            if (this._value != "Off") {
                if (this._noToggleToOff == false) {
                    this.SetChecked(false);
                    this.SetNeedRecalc(true);
                }
            }
            else {
                this.SetChecked(true);
                this.SetNeedRecalc(true);
            }

            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field._exportValue != oThis._exportValue && field._value != "Off") {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
                else if (field._exportValue == oThis._exportValue && oThis._value == "Off") {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
                else if (field._exportValue == oThis._exportValue && field._value == "Off") {
                    field.SetChecked(true);
                    field.SetNeedRecalc(true);
                }
            });

            if (AscCommon.History.Is_LastPointEmpty())
                AscCommon.History.Remove_LastPoint();

        }
        
        private_getViewer()._paintForms();
    };

    /**
	 * Updates all field with this field name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.UpdateAll = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        
        if (this._radiosInUnison) {
            // отмечаем все radiobuttons с тем же экспортом, что и отмеченные
            let sExportValue;
            for (let i = 0; i < aFields.length; i++) {
                if (!sExportValue && aFields[i]._value != "Off") {
                    sExportValue = aFields[i]._exportValue;
                    break;
                }
            }
            if (!sExportValue) {
                aFields.forEach(function(field) {
                    field.SetChecked(false);
                });
            }
            else {
                aFields.forEach(function(field) {
                    if (field._exportValue != sExportValue) {
                        field.SetChecked(false);
                    }
                    else {
                        field.SetChecked(true);
                    }
                });
            }
        }
        else {
            let oCheckedFld = null;
            // оставляем активной первую отмеченную radiobutton
            for (let i = 0; i < aFields.length; i++) {
                if (!oCheckedFld && aFields[i]._value != "Off") {
                    oCheckedFld = aFields[i];
                    continue;
                }
                if (oCheckedFld) {
                    aFields[i].SetChecked(false);
                }
            }
        }
    };
    /**
	 * Set checked to this field (not for all with the same name).
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.SetChecked = function(bChecked) {
        let oRun = this._content.GetElement(0).GetElement(0);
        if (bChecked) {
            oRun.ClearContent();
            oRun.AddText(String.fromCharCode(CheckedSymbol));
            this._value = this._exportValue;
        }
        else {
            oRun.ClearContent();
            oRun.AddText("〇");
            this._value = "Off";
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.ApplyValueForAll = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        for (let i = 0; i < aFields.length; i++) {
            aFields[i].CheckValue();
            aFields[i].SetNeedRecalc(true);
        }
    };
    /**
	 * Checks value of the field and corrects it.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.CheckValue = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        if (oRun.GetText() != "〇")
            this._value = this._exportValue;
        else
            this._value = "Off";
    };
    CRadioButtonField.prototype.SetNoTogleToOff = function(bValue) {
        this._noToggleToOff = bValue;
    };
    CRadioButtonField.prototype.SetRadiosInUnison = function(bValue) {
        this._radiosInUnison = bValue;
    };
    CRadioButtonField.prototype.SetValue = function(sValue) {
        if (sValue != "Off") {
            let oRun = this._content.GetElement(0).GetElement(0);
            oRun.ClearContent();
            oRun.AddText(String.fromCharCode(CheckedSymbol));
            this._value = sValue;
        }
    };

    function CTextField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.text, nPage, aRect);
        
        this._alignment         = ALIGN_TYPE.left;
        this._calcOrderIndex    = 0;
        this._charLimit         = 0; // to do
        this._comb              = false;
        this._defaultStyle      = Object.assign({}, DEFAULT_SPAN); // to do (must not be fileSelect flag)
        this._doNotScroll       = false;
        this._doNotSpellCheck   = false;
        this._multiline         = false;
        this._password          = false;
        this._richText          = false; // to do связанные свойства, методы
        this._richValue         = [];
        this._textFont          = "ArialMT";
        this._fileSelect        = false;

        // internal
        TurnOffHistory();
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
        this._contentFormat = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._contentFormat.ParentPDF = this;
        this._contentFormat.SetUseXLimit(false);

        this._scrollInfo = null;
    }
    CTextField.prototype = Object.create(CBaseField.prototype);
	CTextField.prototype.constructor = CTextField;
    
    CTextField.prototype.SetComb = function(bComb) {
        if (bComb == true) {
            this._comb = true;
            this._doNotScroll = true;
        }
        else {
            this._comb = false;
        }

        this._content.GetElement(0).Content.forEach(function(run) {
            run.RecalcInfo.Measure = true;
        });
        this._contentFormat.GetElement(0).Content.forEach(function(run) {
            run.RecalcInfo.Measure = true;
        });
    };
    CTextField.prototype.SetCharLimit = function(nChars) {
        this._charLimit = nChars;
    };
    CTextField.prototype.SetDoNotScroll = function(bNot) {
        this._doNotScroll = bNot;
    };
    CTextField.prototype.SetDoNotSpellCheck = function(bNot) {
        this._doNotSpellCheck = bNot;
    };
    CTextField.prototype.SetFileSelect = function(bFileSelect) {
        if (bFileSelect === true && this._multiline != true && this._charLimit === 0
            && this.password != true && this.defaultValue == "") {
                this._fileSelect = true;
            }
        else if (bFileSelect === false) {
            this._fileSelect = false;
        }
    };
    CTextField.prototype.SetMultiline = function(bMultiline) {
        if (bMultiline == true && this.fileSelect != true) {
            this._content.SetUseXLimit(true);
            this._contentFormat.SetUseXLimit(true);
            this._multiline = true;
        }
        else if (bMultiline === false) {
            this._content.SetUseXLimit(false);
            this._contentFormat.SetUseXLimit(false);
            this._multiline = false;
        }
    };
    CTextField.prototype.SetPassword = function(bPassword) {
        if (bPassword === true && this.fileSelect != true) {
            this._password = true;
        }
        else if (bPassword === false) {
            this._password = false;
        }
    };
    CTextField.prototype.SetRichText = function(bRichText) {
        this._richText = bRichText;
    };
    CTextField.prototype.SetValue = function(sValue) {
        this._content.GetElement(0).GetElement(0).AddText(sValue);
    };
    CTextField.prototype.SetValueFormat = function(sValue) {
        this._contentFormat.GetElement(0).GetElement(0).AddText(sValue);
    };

    CTextField.prototype.Draw = function(oCtx) {
        let oViewer = private_getViewer();

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = ((this._rect[2]) - (this._rect[0]));
        let nHeight = ((this._rect[3]) - (this._rect[1]));

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        
        this.DrawBorders(oCtx);
        
        let contentX        = (X + nWidth * 0.01) * g_dKoef_pix_to_mm;
        let contentY        = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm;
        let contentXLimit   = (X + nWidth * 0.99) * g_dKoef_pix_to_mm;
        let contentYLimit   = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm;

        let oContentToDraw = this._actions.Format && oViewer.mouseDownFieldObject != this ? this._contentFormat : this._content;

        if ((this.borderStyle == "solid" || this.borderStyle == "dashed") && 
        this._comb == true && this._charLimit > 1) {
            contentX = (X) * g_dKoef_pix_to_mm;
            contentXLimit = (X + nWidth) * g_dKoef_pix_to_mm;
        }
        
        if (this._multiline == false) {
            // выставляем текст посередине
            let nContentH = this._content.GetElement(0).Get_EmptyHeight();
            contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm - nContentH / 2;
        }

        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        this._contentRect.X = contentX;
        this._contentRect.Y = contentY;
        this._contentRect.W = contentXLimit - contentX;
        this._contentRect.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._contentFormat.X = this._oldContentPos.X = contentX;
            this._content.Y      = this._contentFormat.Y = this._oldContentPos.Y = contentY;
            this._content.XLimit = this._contentFormat.XLimit = this._oldContentPos.XLimit = contentXLimit;
            this._content.YLimit = this._contentFormat.YLimit = this._oldContentPos.YLimit = 20000;
            this._content.Recalculate_Page(0, true);
            this._contentFormat.Recalculate_Page(0, true);
        }
        else if (this._needRecalc) {
            oContentToDraw.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.SetNeedRecalc(false);
        }
        
        if (this._multiline == true) {
            oContentToDraw.ResetShiftView();
            oContentToDraw.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        if (this._bAutoShiftContentView)
            this.CheckFormViewWindow();

        let oGraphics = new AscCommon.CGraphics();

        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.AddClipRect(oContentToDraw.X, oContentToDraw.Y, oContentToDraw.XLimit - oContentToDraw.X, contentYLimit - contentY);

        oContentToDraw.Draw(0, oGraphics);

        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && oContentToDraw.IsSelectionUse() == false && oViewer.fieldFillingMode)
            oContentToDraw.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        
        this.SetNeedRecalc(false);
    };
    CTextField.prototype.onMouseDown = function(x, y, e) {
        let {X, Y} = private_getPageCoordsMM(x, y, this._page);

        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1;
        
        this._content.Selection_SetStart(X, Y, 0, e);
        
        this._content.RecalculateCurPos();
        if (this._doNotScroll == false && this._multiline)
            this.UpdateScroll(true);

        if (this.IsNeedDrawFromStream() == true) {
            this.SetDrawFromStream(false);
            private_getViewer()._paintForms();
        }
    };
    CTextField.prototype.SelectionSetStart = function(x, y, e) {
        let {X, Y} = private_getPageCoordsMM(x, y, this._page);
        this._content.Selection_SetStart(X, Y, 0, e);
    };
    CTextField.prototype.SelectionSetEnd = function(x, y, e) {
        let {X, Y} = private_getPageCoordsMM(x, y, this._page);
        this._content.Selection_SetEnd(X, Y, 0, e);
    };
    CTextField.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorLeft(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorRight(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorDown(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorUp(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.EnterText = function(aChars)
    {
        if (aChars.length > 0)
            CreateNewHistoryPointForField(this);

        let isCanEnter = private_doKeystrokeAction(this, aChars);
        if (isCanEnter == false)
            return false;

        let oPara = this._content.GetElement(0);
        if (this._content.IsSelectionUse() && this._content.IsSelectionEmpty())
            this._content.RemoveSelection();

        let nChars = 0;
        function getCharsCount(oRun) {
            var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
                    nChars++;
            }
        }
        
        if (this._content.IsSelectionUse()) {
            // Если у нас что-то заселекчено и мы вводим текст или пробел
			// и т.д., тогда сначала удаляем весь селект.
            this._content.Remove(1, true, false, true);
        }

        if (this._charLimit != 0)
            this._content.CheckRunContent(getCharsCount);

        let nMaxCharsToAdd = this._charLimit != 0 ? this._charLimit - nChars : aChars.length;
        for (let index = 0, count = Math.min(nMaxCharsToAdd, aChars.length); index < count; ++index) {
            let codePoint = aChars[index];
            oPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        if (aChars.length > 0) {
            this.SetNeedRecalc(true);
            this._needApplyToAll = true; // флаг что значение будет применено к остальным формам с таким именем
            this._bAutoShiftContentView = true && this._doNotScroll == false;
        }

        this.SetWasChanged(true);
        return true;
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CTextField
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.ApplyValueForAll = function(bUnionPoints) {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThisPara = this._content.GetElement(0);
        
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        TurnOffHistory();

        if (this._actions.Format) {
            this.SetNeedRecalc(true);
            this._doc.activeForm = this;
            let isValidFormat = eval(this._actions.Format.script);
            // проверка для форматов, не ограниченных на какие-либо символы своей функцией keystroke
            // например для даты, маски
            if (isValidFormat === false && this._apiForm.value != "") {
                // отменяем все изменения сделанные в форме, т.к. не подходят формату 
                this.UnionLastHistoryPoints();
                let nPoint = AscCommon.History.Index;
                AscCommon.History.Undo();
                
                // удаляем точки
                AscCommon.History.Points.length = nPoint;

                // to do выдать предупреждение, что строка не подходит по формату
                return;
            }
        }
        
        if (bUnionPoints)
            this.UnionLastHistoryPoints();

        if (aFields.length == 1)
            this._needApplyToAll = false;

        for (let i = 0; i < aFields.length; i++) {
            aFields[i].SetWasChanged(true); // фиксируем, что форма была изменена

            if (aFields[i].HasShiftView()) {
                aFields[i]._content.MoveCursorToStartPos();

                if (aFields[i] == this) {
                    aFields[i].AddToRedraw();
                    continue;
                }
            }
                
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i].SetNeedRecalc(true);
        }

        let oParaFromFormat = this._contentFormat.GetElement(0);
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._contentFormat.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oParaFromFormat.Content.length - 1; nItem++) {
                oThisRun = oParaFromFormat.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i].SetNeedRecalc(true);
        }
    };
    
    /**
	 * Unions the last history points of this form.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.UnionLastHistoryPoints = function() {
        let oTmpPoint;
        let oResultPoint = {
            State      : undefined,
            Items      : [],
            Time       : new Date().getTime(),
            Additional : {FormFilling: this, CanUnion: false},
            Description: undefined
        };

        let i = 0;
        for (i = AscCommon.History.Points.length - 1; i >= 0 ; i--) {
            oTmpPoint = AscCommon.History.Points[i];
            if (oTmpPoint.Additional.FormFilling === this && oTmpPoint.Additional.CanUnion != false) {
                oResultPoint.Items = oTmpPoint.Items.concat(oResultPoint.Items);
            }
            else {
                break;
            }
        }

        i++; // индекс точки, в которую поместим результирующее значение.
        
        // если больше 2х точек, то заменяем
        if (i < AscCommon.History.Points.length - 1) {
            AscCommon.History.Index = i;
            AscCommon.History.Points.splice(i, AscCommon.History.Points.length - i, oResultPoint);
        }
        else if (AscCommon.History.Points[i])
            AscCommon.History.Points[i].Additional.CanUnion = false; // запрещаем объединять последнюю добавленную точку
    };
    /**
	 * Removes all history points, which were done before form was applied.
	 * @memberof CTextField
     * @param {number} [nCurPoint=AscCommon.History.Index]
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.RemoveNotAppliedChangesPoints = function(nCurPoint) {
        nCurPoint = nCurPoint != undefined ? nCurPoint : AscCommon.History.Index + 1;

        if (!AscCommon.History.Points[nCurPoint + 1] || AscCommon.History.Points[nCurPoint + 1].Additional.CanUnion === false) {
            return;
        }
        AscCommon.History.Points.splice(nCurPoint + 1, AscCommon.History.Points.length - 1);
    };
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Remove = function(nDirection, bWord) {
        CreateNewHistoryPointForField(this);

        this._content.Remove(nDirection, true, false, false, bWord);
        
        if (AscCommon.History.Is_LastPointEmpty())
            AscCommon.History.Remove_LastPoint();
        else {
            this.SetNeedRecalc(true);
            this._needApplyToAll = true;
        }

        this.SetWasChanged(true);
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._alignment         = aFields[i]._alignment;
                this._calcOrderIndex    = aFields[i]._calcOrderIndex;
                this._charLimit         = aFields[i]._charLimit;
                this._comb              = aFields[i]._comb;
                this._doNotScroll       = aFields[i]._doNotScroll;
                this._doNotSpellCheckl  = aFields[i]._doNotSpellCheckl;
                this._fileSelect        = aFields[i]._fileSelect;
                this._multiline         = aFields[i]._multiline;
                this._password          = aFields[i]._password;
                this._richText          = aFields[i]._richText;
                this._richValue         = aFields[i]._richValue.slice();
                this._textFont          = aFields[i]._textFont;
                this._borderStyle       = aFields[i]._borderStyle;

                this._actions = aFields[i]._actions ? aFields[i]._actions.Copy(this) : null;

                if (this._multiline)
                    this._content.SetUseXLimit(true);

                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();

                // format content
                oPara = this._contentFormat.GetElement(0);
                oParaToCopy = aFields[i]._contentFormat.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                break;
            }
        }
    };

    // pdf api methods

    /**
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
     */
    
    function CBaseListField(sName, sType, nPage, aRect)
    {
        CBaseField.call(this, sName, sType, nPage, aRect);

        this._commitOnSelChange     = false;
        this._currentValueIndices   = undefined;
        this._textFont              = "ArialMT";
        this._options               = [];
        
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);
    }
    CBaseListField.prototype = Object.create(CBaseField.prototype);
	CBaseListField.prototype.constructor = CBaseListField;

    /**
	 * Unions the last history points of this form.
	 * @memberof CBaseListField
     * @param {boolean} [bForbidToUnion=true] - wheter to forbid to merge the points united by this iteration
	 * @typeofeditors ["PDF"]
	 */
    CBaseListField.prototype.UnionLastHistoryPoints = function(bForbidToUnion) {
        if (bForbidToUnion == undefined)
            bForbidToUnion = true;
            
        let oTmpPoint;
        let oResultPoint = {
            State      : undefined,
            Items      : [],
            Time       : new Date().getTime(),
            Additional : {FormFilling: this, CanUnion: !bForbidToUnion},
            Description: undefined
        };

        let i = 0;
        for (i = AscCommon.History.Points.length - 1; i >= 0 ; i--) {
            oTmpPoint = AscCommon.History.Points[i];
            if (oTmpPoint.Additional.FormFilling === this && oTmpPoint.Additional.CanUnion != false) {
                oResultPoint.Items = oTmpPoint.Items.concat(oResultPoint.Items);
            }
            else {
                break;
            }
        }

        i++; // индекс точки, в которую поместим результирующее значение.
        
        // если больше 2х точек, то заменяем
        if (i < AscCommon.History.Points.length - 1) {
            AscCommon.History.Index = i;
            AscCommon.History.Points.splice(i, AscCommon.History.Points.length - i, oResultPoint);
        }
        else if (AscCommon.History.Points[i])
            AscCommon.History.Points[i].Additional.CanUnion = !bForbidToUnion; // запрещаем объединять последнюю добавленную точку
    };

    CBaseListField.prototype.SetCommitOnSelChange = function(bValue) {
        this._commitOnSelChange = bValue;
    };

    function CComboBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.combobox, nPage, aRect);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;

        // internal
        this._id = AscCommon.g_oIdCounter.Get_NewId();

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
        this._contentFormat = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._contentFormat.ParentPDF = this;
        this._contentFormat.SetUseXLimit(false);
    };
    CComboBoxField.prototype = Object.create(CBaseListField.prototype);
	CComboBoxField.prototype.constructor = CComboBoxField;

    CComboBoxField.prototype.Draw = function(oCtx) {
        let oViewer = private_getViewer();

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = (this._rect[2] - this._rect[0]);
        let nHeight = (this._rect[3] - this._rect[1]);

        // маркер dropdown
        let nMarkX = X + nWidth * 0.95 + (nWidth * 0.025) - (nWidth * 0.025)/4;
        let nMarkWidth = nWidth * 0.025;
        let nMarkHeight = nMarkWidth/ 2;
        oCtx.beginPath();
        oCtx.moveTo(nMarkX, Y + nHeight/2 + nMarkHeight/2);
        oCtx.lineTo(nMarkX + nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
        oCtx.lineTo(nMarkX - nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
        oCtx.fill();

        this._markRect = {
            x1: (nMarkX - nMarkWidth/2) - ((X + nWidth) - (nMarkX + nMarkWidth/2)),
            y1: Y,
            x2: X + nWidth,
            y2: Y + nHeight
        }

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        this.DrawBorders(oCtx,);
        let oMargins = this.GetBordersWidth();

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        let oContentToDraw = this._actions.Format && oViewer.mouseDownFieldObject != this ? this._contentFormat : this._content;

        // ограничиваем контент позицией маркера
        contentXLimit = this._markRect.x1 * g_dKoef_pix_to_mm; 
        let nContentH = this._content.GetElement(0).Get_EmptyHeight();
        contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm - nContentH / 2;

        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        this._contentRect.X = contentX;
        this._contentRect.Y = contentY;
        this._contentRect.W = contentXLimit - contentX;
        this._contentRect.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._contentFormat.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._contentFormat.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._contentFormat.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._contentFormat.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
            this._contentFormat.Recalculate_Page(0, true);
        }
        else if (this._needRecalc) {
            oContentToDraw.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.SetNeedRecalc(false);
        }
        
        this.CheckFormViewWindow();

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        oGraphics.init(oCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(oContentToDraw.X, oContentToDraw.Y, oContentToDraw.XLimit - oContentToDraw.X, contentYLimit - contentY);

        oContentToDraw.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && oContentToDraw.IsSelectionUse() == false && oViewer.fieldFillingMode)
            oContentToDraw.RecalculateCurPos();
        
        oGraphics.RemoveClip();
    };
    CComboBoxField.prototype.onMouseDown = function(x, y, e) {
        let oViewer = private_getViewer();

        let {X, Y} = private_getPageCoordsMM(x, y, this._page);
        var pageObject = this.getPageByCoords(x - oViewer.x, y - oViewer.y);

        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0; // to do
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1; //

        if (pageObject.x >= this._markRect.x1 && pageObject.x <= this._markRect.x2 && pageObject.y >= this._markRect.y1 && pageObject.y <= this._markRect.y2 && this._options.length != 0) {
            editor.sendEvent("asc_onShowPDFFormsActions", this, x, y);
            this._content.MoveCursorToStartPos();
        }
        else {
            this._content.Selection_SetStart(X, Y, 0, e);
            this._content.RemoveSelection();
        }
        
        this._content.RecalculateCurPos();
    };
    
    /**
	 * Selects the specified option.
	 * @memberof CComboBoxField
     * @param {boolean} [bAddToHistory=true] - whether to add change to history.
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SelectOption = function(nIdx, bAddToHistory) {
        if (bAddToHistory == undefined)
            bAddToHistory = true;

        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);

        this._currentValueIndices = nIdx;

        if (bAddToHistory) {
            CreateNewHistoryPointForField(this);
        }
        else
            TurnOffHistory();

        oRun.ClearContent();
        if (Array.isArray(this._options[nIdx])) {
            oRun.AddText(this._options[nIdx][0]);
            this._value = this._options[nIdx][0];
        }
        else {
            oRun.AddText(this._options[nIdx]);
            this._value = this._options[nIdx];
        }

        this.SetNeedRecalc(true);
        this._needApplyToAll = true;
    };

    CComboBoxField.prototype.SetValue = function(sValue) {
        let sTextToAdd = "";
        for (let i = 0; i < this._options.length; i++) {
            if (this._options[i][1] && this._options[i][1] == sValue) {
                sTextToAdd = this._options[i][0];
                break;
            }
        }
        if (sTextToAdd == "") {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i] == sValue) {
                    sTextToAdd = this._options[i];
                    break;
                }
            }
        }
        
        if (sTextToAdd == "")
            sTextToAdd = sValue;

        this._content.GetElement(0).GetElement(0).AddText(sTextToAdd);
        this.CheckCurValueIndex();
    };
    CComboBoxField.prototype.SetValueFormat = function(sValue) {
        this._contentFormat.GetElement(0).GetElement(0).AddText(sValue);
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {

                this._calcOrderIndex    = aFields[i]._calcOrderIndex;
                this._doNotSpellCheck   = aFields[i]._doNotSpellCheck;
                this._editable          = aFields[i]._editable;

                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                this._options = aFields[i]._options.slice();
                break;
            }
        }
    };
    CComboBoxField.prototype.EnterText = function(aChars, bForce)
    {
        if (this._editable == false && !bForce)
            return false;

        if (aChars.length > 0)
            CreateNewHistoryPointForField(this);
        else
            return false;

        let isCanEnter = private_doKeystrokeAction(this, aChars);
        if (isCanEnter == false)
            return false;

        let oPara = this._content.GetElement(0);
        if (this._content.IsSelectionUse()) {
            // Если у нас что-то заселекчено и мы вводим текст или пробел
			// и т.д., тогда сначала удаляем весь селект.
            this._content.Remove(1, true, false, true);
        }
        
        for (let index = 0, count = aChars.length; index < count; ++index) {
            let codePoint = aChars[index];
            oPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        this.CheckCurValueIndex();
        this.SetNeedRecalc(true);
        this._needApplyToAll = true; // флаг что значение будет применено к остальным формам с таким именем
        
        return true;
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CComboBoxField
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.ApplyValueForAll = function(bUnionPoints) {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThisPara = this._content.GetElement(0);
        
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        TurnOffHistory();

        if (this._actions.Format) {
            this.SetNeedRecalc(true);
            this._doc.activeForm = this;
            let isValidFormat = eval(this._actions.Format.script);
            // проверка для форматов, не ограниченных на какие-либо символы своей функцией keystroke
            // например для даты, маски
            if (isValidFormat === false && this._apiForm.value != "") {
                // отменяем все изменения сделанные в форме, т.к. не подходят формату 
                this.UnionLastHistoryPoints();
                let nPoint = AscCommon.History.Index;
                AscCommon.History.Undo();
                
                // удаляем точки
                AscCommon.History.Points.length = nPoint;

                // to do выдать предупреждение, что строка не подходит по формату
                return;
            }
        }

        this.CheckCurValueIndex();

        if (bUnionPoints)
            this.UnionLastHistoryPoints(true);

        if (aFields.length == 1)
            this._needApplyToAll = false;

        for (let i = 0; i < aFields.length; i++) {
            if (this.HasShiftView())
                aFields[i]._content.MoveCursorToStartPos();

            if (aFields[i] == this) {
                if (this.HasShiftView())
                    this.AddToRedraw();

                continue;
            }

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i]._currentValueIndices = this._currentValueIndices;
            aFields[i].SetNeedRecalc(true);
        }
    };
    /**
	 * Checks curValueIndex, corrects it and return.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CComboBoxField.prototype.CheckCurValueIndex = function() {
        let sValue = this._content.GetElement(0).GetText({ParaEndToSpace: false});
        this._value = sValue;
        let nIdx = -1;
        if (Array.isArray(this._options) == true) {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i][0] === sValue) {
                    nIdx = i;
                    break;
                }
            }
        }
        else {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i] === sValue) {
                    nIdx = i;
                    break;
                }
            }
        }

        this._currentValueIndices = nIdx;
        return nIdx;
    };

    CComboBoxField.prototype.SetEditable = function(bValue) {
        this._editable = bValue;
    };
    CComboBoxField.prototype.SetOptions = function(aOpt) {
        let aOptToPush = [];
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "")
                aOptToPush.push(aOpt[i]);
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    aOptToPush.push([aOpt[i][0].toString(), aOpt[i][1].toString()])
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                aOptToPush.push(aOpt[i].toString());
            }
        }

        this._options = aOptToPush;
    };
    
    function CListBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.listbox, nPage, aRect);

        this._multipleSelection = false;

        // internal
        this._scrollInfo = null;
        this._bAutoShiftContentView = true;
    };
    CListBoxField.scrollCount = 0;
    CListBoxField.prototype = Object.create(CBaseListField.prototype);
	CListBoxField.prototype.constructor = CListBoxField;

    CListBoxField.prototype.Draw = function(oCtx) {
        let oViewer = private_getViewer();

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = (this._rect[2] - this._rect[0]);
        let nHeight = (this._rect[3] - this._rect[1]);

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        this.DrawBorders(oCtx);
        let oMargins = this.GetBordersWidth();

        let contentX = (X + nWidth * 0.01 + oMargins.left) * g_dKoef_pix_to_mm;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit = (X + nWidth * 0.99 - oMargins.right) * g_dKoef_pix_to_mm;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;

        this._contentRect.X = contentX;
        this._contentRect.Y = contentY;
        this._contentRect.W = contentXLimit - contentX;
        this._contentRect.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (this._needRecalc) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.SetNeedRecalc(false);
        }
        
        if (this._bAutoShiftContentView)
            this.CheckFormViewWindow();
        else {
            this._content.ResetShiftView();
            this._content.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;
        oGraphics.init(oCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        
        oGraphics.RemoveClip();
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._multipleSelection = aFields[i]._multipleSelection;
                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++) {
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                }
                
                this._options = aFields[i]._options.slice();
                this._currentValueIndices = aFields.multipleSelection ? aFields[i]._currentValueIndices.slice() : aFields[i]._currentValueIndices;

                let oPara;
                for (let i = 0; i < this._content.Content.length; i++) {
                    oPara = this._content.GetElement(i);
                    if (oPara.Pr.Shd && oPara.Pr.Shd.IsNil() == false)
                        oPara.private_CompileParaPr(true);
                }
                break;
            }
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CListBoxField
     * @param {CListBoxField} [oFieldToSkip] - field to don't be apply changes.
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.ApplyValueForAll = function(oFieldToSkip, bUnionPoints) {
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        let aFields = this._doc.getWidgetsByName(this.GetFullName());
        let oThis = this;
        
        
        if (bUnionPoints)
            this.UnionLastHistoryPoints();

        this.CheckFormViewWindow();
        this.RevertContentViewToOriginal();

        let oThisBounds = this.getFormRelRect();

        aFields.forEach(function(field) {
            field.SetWasChanged(true);

            field._bAutoShiftContentView = false;

            field.SetNeedRecalc(true);
            if (field.HasShiftView()) {
                if (field == oThis) {
                    field.AddToRedraw();
                    return;
                }
            }
            if (oThis == field || oFieldToSkip == field)
                return;

            if (oThis._multipleSelection) {
                // снимаем выделение с тех, которые не присутсвуют в поле, от которого применяем ко всем
                for (let i = 0; i < field._currentValueIndices.length; i++) {
                    if (oThis._currentValueIndices.includes(field._currentValueIndices[i]) == false) {
                        field.UnselectOption(field._currentValueIndices[i]);
                    }
                }
                
                for (let i = 0; i < oThis._currentValueIndices.length; i++) {
                    // добавляем выделение тем, которые не присутсвуют в текущем поле, но присутсвуют в том, от которого применяем
                    if (field._currentValueIndices.includes(oThis._currentValueIndices[i]) == false) {
                        field.SelectOption(oThis._currentValueIndices[i], false, false);
                    }
                }
                field._currentValueIndices = oThis._currentValueIndices.slice();
            }
            else {
                field._currentValueIndices = oThis._currentValueIndices;
                field.SelectOption(field._currentValueIndices, true, false);
            }

            let oFieldBounds = field.getFormRelRect();
            if (Math.abs(oFieldBounds.H - oThisBounds.H) > 0.001) {
                field.CheckFormViewWindow();
                field.RevertContentViewToOriginal();
            }
            else {
                field._curShiftView.x = oThis._curShiftView.x;
                field._curShiftView.y = oThis._curShiftView.y;
                field._originShiftView.x = oThis._originShiftView.x;
                field._originShiftView.y = oThis._originShiftView.y;
            }
        });
    };
    
    CListBoxField.prototype.SetMultipleSelection = function(bValue) {
        if (bValue == true) {
            this._multipleSelection = true;
            this._currentValueIndices = [this._currentValueIndices];
        }
        else {
            this._multipleSelection = false;
            this._currentValueIndices = this._currentValueIndices[0];
            this._currentValueIndices != -1 && this.SelectOption(this._currentValueIndices, true);
        }
    };

    CListBoxField.prototype.SelectOption = function(nIdx, isSingleSelect, bAddToHistory) {
        if (bAddToHistory == undefined)
            bAddToHistory = true;

        let oPara = this._content.GetElement(nIdx);
        let oApiPara;
        
        if (bAddToHistory) {
            CreateNewHistoryPointForField(this);
        }
        else
            TurnOffHistory();

        this._content.Set_CurrentElement(nIdx);
        if (isSingleSelect) {
            this._content.Content.forEach(function(para) {
                oApiPara = editor.private_CreateApiParagraph(para);
                if (para.Pr.Shd && para.Pr.Shd.IsNil() == false) {
                    oApiPara.SetShd('nil');
                    para.private_CompileParaPr(true);
                }
            });
        }

        if (oPara) {
            oApiPara = editor.private_CreateApiParagraph(oPara);
            oApiPara.SetShd('clear', LISTBOX_SELECTED_COLOR.r, LISTBOX_SELECTED_COLOR.g, LISTBOX_SELECTED_COLOR.b);
            oApiPara.Paragraph.private_CompileParaPr(true);
        }

        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
        this._needApplyToAll = true;

        if (this.IsNeedDrawFromStream() == true) {
            this.SetDrawFromStream(false);
            private_getViewer()._paintForms();
        }
    };
    CListBoxField.prototype.UnselectOption = function(nIdx) {
        let oApiPara = editor.private_CreateApiParagraph(this._content.GetElement(nIdx));
        oApiPara.SetShd('nil');
        oApiPara.Paragraph.private_CompileParaPr(true);
        this.SetNeedRecalc(true);
    };
    CListBoxField.prototype.SetOptions = function(aOpt) {
        this._content.Internal_Content_RemoveAll();
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            sCaption = "";
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "") {
                sCaption = aOpt[i];
                this._options.push(aOpt[i]);
            }
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    this._options.push([aOpt[i][0].toString(), aOpt[i][1].toString()]);
                    sCaption = aOpt[i][0].toString();
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                this._options.push(aOpt[i].toString());
                sCaption = aOpt[i].toString();
            }

            if (sCaption != "") {
                oPara = new AscCommonWord.Paragraph(this._content.DrawingDocument, this._content, false);
                oRun = new AscCommonWord.ParaRun(oPara, false);
                this._content.Internal_Content_Add(i, oPara);
                oPara.Add(oRun);
                oRun.AddText(sCaption);
            }
        }
    };
    CListBoxField.prototype.SetValue = function(value) {
        let aIndexes = [];
        if (Array.isArray(value)) {
            for (let sVal of value) {
                let isFound = false;
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i][1] && this._options[i][1] == sVal) {
                        if (aIndexes.includes(i))
                            continue;
                        else {
                            isFound = true;
                            aIndexes.push(i);
                            break;
                        }
                    }
                }
                if (isFound == false) {
                    for (let i = 0; i < this._options.length; i++) {
                        if (this._options[i] == sVal) {
                            if (aIndexes.includes(i))
                                continue;
                            else {
                                aIndexes.push(i);
                                break;
                            }
                        }
                    }
                }
            }
        }
        else {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i][1] && this._options[i][1] == value) {
                    aIndexes.push(i);
                    break;
                }
            }
            if (aIndexes.length == 0) {
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i] == value) {
                        aIndexes.push(i);
                        break;
                    }
                }
            }
        }
        
        let oPara, oApiPara;
        for (let idx of aIndexes) {
            oPara = this._content.GetElement(idx);
            oApiPara = editor.private_CreateApiParagraph(oPara);
            oApiPara.SetShd('clear', LISTBOX_SELECTED_COLOR.r, LISTBOX_SELECTED_COLOR.g, LISTBOX_SELECTED_COLOR.b);
            oPara.private_CompileParaPr(true);
        }

        this._currentValueIndices = aIndexes.length > 1 ? aIndexes : aIndexes[0];
    };
    CListBoxField.prototype.InsertOption = function(sName, sExport, nIdx) {
        let optToInsert = sExport ? [sName, sExport] : sName;
        if (nIdx == -1 || nIdx > this._options.length) {
            nIdx = this._options.length;
        }

        let oIdxItem = this._content.GetElement(nIdx);

        this._options = this._options.splice(nIdx, 0, optToInsert);
        
        let oPara = new AscCommonWord.Paragraph(this._content.DrawingDocument, this._content, false);
        let oRun = new AscCommonWord.ParaRun(oPara, false);
        this._content.Internal_Content_Add(nIdx, oPara);
        oPara.Add(oRun);
        oRun.AddText(sName);
    };

    CListBoxField.prototype.onMouseDown = function(x, y, e) {
        if (this._options.length == 0)
            return;

        let oViewer = private_getViewer();
        let {X, Y} = private_getPageCoordsMM(x, y, this._page);
        
        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1;
        
        let nPos = this._content.Internal_GetContentPosByXY(X, Y, 0);

        if (this._multipleSelection == true) {
            if (e.ctrlKey == true) {
                if (this._currentValueIndices.includes(nPos)) {
                    this.UnselectOption(nPos);
                    this._currentValueIndices.splice(this._currentValueIndices.indexOf(nPos), 1);
                }
                else {
                    this.SelectOption(nPos, false);
                    this._currentValueIndices.push(nPos);
                    this._currentValueIndices.sort();
                }
            }
            else {
                this.SelectOption(nPos, true);
                this._currentValueIndices = [nPos];
            }
            this.AddToRedraw();
        }
        else {
            if (nPos == this._currentValueIndices) {
                this.UpdateScroll(true);
                return;
            }
                
            this.SelectOption(nPos, true);
            this._currentValueIndices = nPos;
            this.AddToRedraw();
        }

        this._bAutoShiftContentView = true;
        this.UnionLastHistoryPoints(false);

        oViewer._paintForms();
    };
    CListBoxField.prototype.MoveSelectDown = function() {
        this._bAutoShiftContentView = true;
        this._content.MoveCursorDown();

        if (this._multipleSelection == true) {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }
        
        this.AddToRedraw();
        private_getViewer()._paintForms();
        this.UpdateScroll(true);
        this
    };
    CListBoxField.prototype.MoveSelectUp = function() {
        this._bAutoShiftContentView = true;
        this._content.MoveCursorUp();

        if (this._multipleSelection == true) {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }

        this.AddToRedraw();
        private_getViewer()._paintForms();
        this.UpdateScroll(true);
    };
    CListBoxField.prototype.UpdateScroll = function(bShow) {
        let oContentBounds  = this._content.GetContentBounds(0);
        let oContentRect    = this.getFormRelRect();
        let oFormRect       = this.getFormRect();

        let nContentH       = oContentBounds.Bottom - oContentBounds.Top;
        let oScroll, oScrollDocElm, oScrollSettings;

        if (nContentH < oContentRect.H)
            return;

        if (typeof(bShow) != "boolean" && this._scrollInfo)
            bShow = this._scrollInfo.scroll.canvas.style.display == "none" ? false : true;

        if (bShow == false) {
            if (this._scrollInfo)
                this._scrollInfo.scroll.canvas.style.display = "none";
            return;
        }

        let oViewer = private_getViewer();
        
        let oGlobalCoords1  = private_getGlobalCoordsByPageCoords(oFormRect.X, oFormRect.Y, this._page);
        let oGlobalCoords2  = private_getGlobalCoordsByPageCoords(oFormRect.X + oFormRect.W, oFormRect.Y + oFormRect.H, this._page);
        let oBorderWidth    = this.GetBordersWidth(true);
        
        if (this._scrollInfo == null && oContentBounds.Bottom - oContentBounds.Top > oContentRect.H) {
            CListBoxField.scrollCount++;
            oScrollDocElm = document.createElement('div');
            document.getElementById('editor_sdk').appendChild(oScrollDocElm);
            oScrollDocElm.id = "formScroll_" + CListBoxField.scrollCount;
            oScrollDocElm.style.top = oGlobalCoords1.Y - oBorderWidth.top / 2 + 'px';
            oScrollDocElm.style.left = oGlobalCoords2.X + 'px';
            oScrollDocElm.style.position = "absolute";
            oScrollDocElm.style.display = "block";
			oScrollDocElm.style.width = "14px";
			oScrollDocElm.style.height = (oGlobalCoords2.Y - oGlobalCoords1.Y) + (oBorderWidth.top + oBorderWidth.bottom) / 2 + "px";

            let nMaxShiftY = oContentRect.H - nContentH;

            oScrollSettings = editor.WordControl.CreateScrollSettings();
            oScrollSettings.isHorizontalScroll = false;
		    oScrollSettings.isVerticalScroll = true;
		    oScrollSettings.contentH = Math.abs(nMaxShiftY);
            oScrollSettings.screenH = 0;
            oScrollSettings.scrollerMinHeight = 5;
            
            let nScrollCoeff = this._curShiftView.y / nMaxShiftY;
            
            let oPara = this._content.GetElement(0);
            let oCurParaHeight  = oPara.Lines[0].Bottom - oPara.Lines[0].Top;

            oScrollSettings.vscrollStep = oCurParaHeight;
            oScroll = new AscCommon.ScrollObject(oScrollDocElm.id, oScrollSettings);

            let oThis = this;
            oScroll.bind("scrollvertical", function(evt) {
                oThis.ScrollVertical(evt.scrollD, evt.maxScrollY);
            });
            oScroll.bind("mouseup", function(evt) {
                if (oThis.type == "listbox")
                    oThis.ScrollVerticalEnd();
            });

            oScroll.scrollVCurrentY = oScroll.maxScrollY * nScrollCoeff;
            
            this._scrollInfo = {
                scroll:         oScroll,
                docElem:        oScrollDocElm,
                baseYPos:       parseInt(oScrollDocElm.style.top),
                oldZoom:        oViewer.zoom,
                scrollCoeff:    nScrollCoeff // проскроленная часть
            }

            oScroll.Repos(oScrollSettings, false);
        }
        else if (this._scrollInfo) {
            let nMaxShiftY = oContentRect.H - nContentH;
            let needUpdatePos = this._scrollInfo.oldZoom != oViewer.zoom || oGlobalCoords1.Y - oBorderWidth.top != this._scrollInfo.baseYPos;

            if (needUpdatePos) {
                oScrollSettings = editor.WordControl.CreateScrollSettings();
                oScrollSettings.isHorizontalScroll = false;
                oScrollSettings.isVerticalScroll = true;
                oScrollSettings.contentH = Math.abs(nMaxShiftY);
                oScrollSettings.screenH = 0;
                oScrollSettings.scrollerMinHeight = 5;
                this._scrollInfo.scroll.scrollVCurrentY = this._scrollInfo.scroll.maxScrollY * this._scrollInfo.scrollCoeff;
                this._scrollInfo.scroll.Repos(oScrollSettings, false);
                let nScrollCoeff = this._content.ShiftViewY / nMaxShiftY;
                this._scrollInfo.scrollCoeff = nScrollCoeff;

                this._scrollInfo.docElem.style.top      = oGlobalCoords1.Y - oBorderWidth.top / 2 + 'px';
                this._scrollInfo.docElem.style.left     = oGlobalCoords2.X + 'px';
                this._scrollInfo.docElem.style.height   = (oGlobalCoords2.Y - oGlobalCoords1.Y) + (oBorderWidth.top + oBorderWidth.bottom) / 2 + "px";
            
                this._scrollInfo.oldZoom = oViewer.zoom;
                this._scrollInfo.baseYPos = parseInt(this._scrollInfo.docElem.style.top);
            }

            if (bShow === true)
                this._scrollInfo.scroll.canvas.style.display = "";
            if (bShow === false)
                this._scrollInfo.scroll.canvas.style.display = "none";
        }
    };
    CListBoxField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        this._bAutoShiftContentView = false;

        let nScrollCoeff                = scrollY / maxYscroll;
        this._curShiftView.y            = -nScrollCoeff * maxYscroll;
        this._scrollInfo.scrollCoeff    = nScrollCoeff;
        this.AddToRedraw();
        private_getViewer()._paintForms();
    };
    CListBoxField.prototype.ScrollVerticalEnd = function() {
        let nHeightPerPara  = this._content.GetElement(1).Y - this._content.GetElement(0).Y;
        let nShiftCount     = this._curShiftView.y / nHeightPerPara; // количество смещений в длинах параграфов
        if (Math.abs(Math.round(nShiftCount) - nShiftCount) <= 0.001)
            return;

        let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
        this._curShiftView.y            = Math.round(nShiftCount) * nHeightPerPara;
        this._bAutoShiftContentView     = false;
        this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        
        this.AddToRedraw();
        private_getViewer()._paintForms();
    };
    /**
	 * Checks curValueIndices, corrects it and return.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CListBoxField.prototype.CheckCurValueIndex = function() {
        let nIdx;
        if (this._multipleSelection)
            nIdx = []
        else
            nIdx = -1;

        let oCurPara;
        for (let i = 0; i < this._content.Content.length; i++) {
            oCurPara = this._content.GetElement(i);
            if (oCurPara.Pr.Shd && oCurPara.Pr.Shd.IsNil() == false) {
                if (this._multipleSelection)
                    nIdx.push(i);
                else {
                    nIdx = i;
                    break;
                }
            }
        }

        this._currentValueIndices = nIdx;
        return nIdx;
    };
    CListBoxField.prototype.CheckFormViewWindow = function()
    {
        let nFirstSelectedPara = Array.isArray(this._currentValueIndices) ? this._currentValueIndices[0] : this._currentValueIndices;
        let oParagraph  = this._content.GetElement(nFirstSelectedPara);

        // размеры всего контента
        let oPageBounds     = this._content.GetContentBounds(0);
        let oCurParaHeight  = oParagraph.Lines[0].Bottom - oParagraph.Lines[0].Top;

        let oFormBounds = this.getFormRelRect();
        let nDy = 0, nDx = 0;
        
        if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
            if (oParagraph.Y + oCurParaHeight > oFormBounds.Y + oFormBounds.H)
                nDy = oFormBounds.Y + oFormBounds.H - (oParagraph.Y + oCurParaHeight);
            else if (oParagraph.Y < oFormBounds.Y)
                nDy = oFormBounds.Y - oParagraph.Y;
            else if (oParagraph.Y + oCurParaHeight < oFormBounds.Y)
                nDy = oCurParaHeight;
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this._content.ShiftView(nDx, nDy);
            this._originShiftView = {
                x: this._content.ShiftViewX,
                y: this._content.ShiftViewY
            }
            
            this._curShiftView.x = this._originShiftView.x;
            this._curShiftView.y = this._originShiftView.y;
        }
        else {
            this._originShiftView.x = this._curShiftView.x;
            this._originShiftView.y = this._curShiftView.y;
        }
    };

    function CSignatureField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.signature, nPage, aRect);
    };

    function CFormActions() {
        this.MouseUp = null; 
        this.MouseDown = null; 
        this.MouseEnter = null; 
        this.MouseExit = null; 
        this.OnFocus = null; 
        this.OnBlur = null; 
        this.Keystroke = null; 
        this.Validate = null; 
        this.Calculate = null; 
        this.Format = null;
    }
    CFormActions.prototype.Copy = function() {
        let newObj = new CFormActions();
        if (this.MouseUp != null)
            newObj.MouseUp = this.MouseUp.Copy(); 
        if (this.MouseDown != null)
            newObj.MouseDown = this.MouseDown.Copy(); 
        if (this.MouseEnter != null)
            newObj.MouseEnter = this.MouseEnter.Copy(); 
        if (this.MouseExit != null)
            newObj.MouseExit = this.MouseExit.Copy(); 
        if (this.OnFocus != null)
            newObj.OnFocus = this.OnFocus.Copy(); 
        if (this.OnBlur != null)
            newObj.OnBlur = this.OnBlur.Copy(); 
        if (this.Keystroke != null)
            newObj.Keystroke = this.Keystroke.Copy(); 
        if (this.Validate != null)
            newObj.Validate = this.Validate.Copy(); 
        if (this.Calculate != null)
            newObj.Calculate = this.Calculate.Copy(); 
        if (this.Format != null)
            newObj.Format = this.Format.Copy();

        return newObj;
    }

    function CFormAction(type, sScript) {
        this.type = type;
        this.script = sScript;
    }
    CFormAction.prototype.Copy = function() {
        return new CFormAction(this.type, this.script);
    }
    
    CBaseField.prototype.CheckFormViewWindow = function()
    {
        let oParagraph  = this._content.GetElement(this._content.CurPos.ContentPos);
        let nEndPos     = this._content.Pages[0].EndPos;

        // размеры до текущего параграффа
        this._content.Pages[0].EndPos = this._content.CurPos.ContentPos;
        let oPageBoundsToCurPara = this._content.GetContentBounds(0);
        this._content.Pages[0].EndPos = nEndPos;

        // размеры всего контента
        let oPageBounds     = this._content.GetContentBounds(0);
        let oCurParaHeight  = oParagraph.Lines[0].Bottom - oParagraph.Lines[0].Top;

        let oFormBounds = this.getFormRelRect();

        let nDx = 0, nDy = 0;

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oPageBounds.Left > oFormBounds.X)
                nDx = -oPageBounds.Left + oFormBounds.X;
            else if (oPageBounds.Right < oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oPageBounds.Right;
        }
        else
        {
            nDx = -this._content.ShiftViewX;
        }

        // если высота контента больше чем высота формы (для нескольких параграфов)
        if (this.type == "listbox") {
            if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                if (oPageBoundsToCurPara.Bottom > oFormBounds.Y + oFormBounds.H)
                    nDy = oFormBounds.Y + oFormBounds.H - oPageBoundsToCurPara.Bottom;
                else if (oPageBoundsToCurPara.Bottom - oCurParaHeight < oFormBounds.Y)
                    nDy = oFormBounds.Y - (oPageBoundsToCurPara.Bottom - oCurParaHeight);
                else if (oPageBoundsToCurPara.Bottom < oFormBounds.Y)
                    nDy = oCurParaHeight;
            }
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this._content.ShiftView(nDx, nDy);
            this._originShiftView = {
                x: this._content.ShiftViewX,
                y: this._content.ShiftViewY
            }
        }

        var oCursorPos  = oParagraph.GetCalculatedCurPosXY();
        var oLineBounds = oParagraph.GetLineBounds(oCursorPos.Internal.Line);
        var oLastLineBounds = oParagraph.GetLineBounds(oParagraph.GetLinesCount() - 1);

	    nDx = 0;
	    nDy = 0;

        var nCursorT = Math.min(oCursorPos.Y, oLineBounds.Top);
        var nCursorB = Math.max(oCursorPos.Y + oCursorPos.Height, oLineBounds.Bottom);
        var nCursorH = Math.max(0, nCursorB - nCursorT);

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oCursorPos.X < oFormBounds.X)
                nDx = oFormBounds.X - oCursorPos.X;
            else if (oCursorPos.X > oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oCursorPos.X;
        }

        if (this._multiline) {
            // если высота контента больше чем высота формы
            if (oParagraph.IsSelectionUse()) {
                if (oParagraph.GetSelectDirection() == 1) {
                    if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else {
                    if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                }
            }
            else {
                if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                    if (oLastLineBounds.Bottom - Math.floor(((oFormBounds.Y + oFormBounds.H) * 1000)) / 1000 < 0)
                        nDy = oFormBounds.Y + oFormBounds.H - oLastLineBounds.Bottom;
                    else if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                    else if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else
                    nDy = -this._content.ShiftViewY;
            }
        }

        if (this.type == "listbox") nDx = 0;

        if (Math.abs(nDx) > 0.001 || Math.round(nDy) != 0)
        {
            this._content.ShiftView(nDx, nDy);
            this._curShiftView = {
                x: this._content.ShiftViewX,
                y: this._content.ShiftViewY
            }
        }
    };
    CBaseField.prototype.RevertContentViewToOriginal = function() {
        this._content.ResetShiftView();
        this._curShiftView.x = this._originShiftView.x;
        this._curShiftView.y = this._originShiftView.y;

        this._bAutoShiftContentView = false;
        this._content.ShiftView(this._originShiftView.x, this._originShiftView.y);

        if (this._scrollInfo) {
            let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
            this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        }
    };
    CBaseField.prototype.IsNeedRevertShiftView = function() {
        if (this._curShiftView.y != this._originShiftView.y ||
            this._curShiftView.x != this._originShiftView.x)
            return true;
    };
    CBaseField.prototype.GetBordersWidth = function(bZoomed) {
        let oViewer = private_getViewer();
        let nLineWidth = bZoomed == true ? this._lineWidth * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom : this._lineWidth * AscCommon.AscBrowser.retinaPixelRatio;

        if (nLineWidth == 0 || this.type == "radiobutton") {
            return {
                left:     0,
                top:      0,
                right:    0,
                bottom:   0
            }
        }

        switch (this._borderStyle) {
            case BORDER_TYPES.solid:
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case BORDER_TYPES.beveled:
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }             
            case BORDER_TYPES.dashed:
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case BORDER_TYPES.inset:
                return {
                    left:     2 * nLineWidth,
                    top:      2 * nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }
            case BORDER_TYPES.underline:
                return {
                    left:     0,
                    top:      0,
                    right:    0,
                    bottom:   nLineWidth
                }
        }
    };
    CBaseField.prototype.HasShiftView = function() {
        if (this._content.ShiftViewX != 0 || this._content.ShiftViewY != 0)
            return true;

        return false;
    };
    CBaseField.prototype.MoveCursorToStartPos = function() {
        this._content.MoveCursorToStartPos();
    };
    CBaseField.prototype.SetAlign = function(nAlgnType) {
        this._content.SetApplyToAll(true);
        this._content.SetParagraphAlign(nAlgnType);
        this._content.GetElement(0).private_CompileParaPr(true);
        this._content.SetApplyToAll(false);

        if (this._contentFormat) {
            this._contentFormat.SetApplyToAll(true);
            this._contentFormat.SetParagraphAlign(nAlgnType);
            this._contentFormat.SetApplyToAll(false);
            this._contentFormat.GetElement(0).private_CompileParaPr(true);
        }
    };
    CBaseField.prototype.SetBorderType = function(nType) {
        this._borderStyle = nType;
    };
    CBaseField.prototype.SetBorderWidth = function(nType) {
        this._borderWidth = nType;
        this._lineWidth = nType;
    };
    /**
     * Returns a canvas with origin view (from appearance stream) of current form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CBaseField.prototype.GetOriginView = function() {
        if (this._apIdx == -1)
            return null;

        let oViewer = private_getViewer();
        let oFile   = oViewer.file;
        let oPage   = oViewer.drawingPages[this._page];

        let w = (oPage.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h = (oPage.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

        let aFormsInfo      = oFile.nativeFile.getInteractiveFormsAP(this._page, w, h);
        let oApearanceInfo  = aFormsInfo[this._apIdx];

        let canvas  = document.createElement("canvas");
        let nWidth  = oApearanceInfo["w"];
        let nHeight = oApearanceInfo["h"];

        canvas.width    = nWidth;
        canvas.height   = nHeight;

        canvas.x    = oApearanceInfo["x"];
        canvas.y    = oApearanceInfo["y"];
        
        let supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;

        let ctx             = canvas.getContext("2d");
        let mappedBuffer    = new Uint8ClampedArray(oFile.memory().buffer, oApearanceInfo["retValue"], 4 * nWidth * nHeight);
        let imageData       = null;

        if (supportImageDataConstructor)
        {
            imageData = new ImageData(mappedBuffer, nWidth, nHeight);
        }
        else
        {
            imageData = ctx.createImageData(nWidth, nHeight);
            imageData.data.set(mappedBuffer, 0);                    
        }
        if (ctx)
            ctx.putImageData(imageData, 0, 0);
        
        oViewer.file.free(oApearanceInfo["retValue"]);

        return canvas;
    };
    CBaseField.prototype.DrawOriginView = function(oCtx) {
        let oViewer = private_getViewer();
        let originView = this.GetOriginView();
        let nScale = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;

        let X = this._rect[0];
        let Y = this._rect[1];
        let nWidth = ((this._rect[2]) - (this._rect[0]));
        let nHeight = ((this._rect[3]) - (this._rect[1]));

        if (originView) {
            oCtx.clearRect(X * nScale, Y * nScale, (nWidth + 1) * nScale, ((nHeight + 1) * nScale));

            let x = Math.round((X * nScale));
            let y = Math.round((Y * nScale));
            let w = Math.round(((nWidth + 1) * nScale));
            let h = Math.round(((nHeight + 1) * nScale));

            oCtx.drawImage(originView, 0, 0, originView.width, originView.height, originView.x, originView.y, originView.width, originView.height);
            return;
        }
    };
    // for format

    /**
	 * Convert field value to specific number format.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Format(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;

        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }

        oCultureInfo.NumberGroupSizes = [3];
        
        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm._apiForm.value;
        if (sCurValue == "") {
            oTargetRun.ClearContent();
            return;
        }
            
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;

        if (bCurrencyPrepend)
            sRes = strCurrency + sRes;
        else
            sRes = sRes + strCurrency;

        if (sRes.indexOf("-") != - 1) {
            sRes = sRes.replace("-", "");
            switch (negStyle) {
                case 0:
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    break;
                case 1:
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    break;
                case 2:
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    sRes = "(" + sRes + ")";
                    break;
                case 3:
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    sRes = "(" + sRes + ")";
                    break;
            }
        }
        else {
            oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
        }
        
        oTargetRun.RecalcInfo.TextPr = true
        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Keystroke(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sResultText.indexOf(",") != -1)
                    return false;

                if (isValidNumber(sResultText) == false)
                    return false;
                break;
            case 2:
            case 3:
                if (sResultText.indexOf(".") != -1)
                    return false;

                sResultText = sResultText.replace(/\,/g, ".");
                if (isValidNumber(sResultText) == false)
                    return false;
                break;
        }

        return true;
    }

    /**
	 * Convert field value to specific percent format.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Format(nDec, sepStyle) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;

        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }
        oCultureInfo.NumberGroupSizes = [3];

        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm._apiForm.value;
        sCurValue.replace(",", ".");
        if (sCurValue == "")
            sCurValue = 0;
            
        sCurValue = (parseFloat(sCurValue) * 100).toString();
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;
        sRes = sRes + "%";

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Keystroke(nDec, sepStyle) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sResultText.indexOf(",") != -1)
                    return false;

                if (isValidNumber(sResultText) == false)
                    return false;
                break;
            case 2:
            case 3:
                if (sResultText.indexOf(".") != -1)
                    return false;

                sResultText = sResultText.replace(/\,/g, ".");
                if (isValidNumber(sResultText) == false)
                    return false;
                break;
        }

        return true;
    }

    /**
	 * Convert field value to specific date format.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Format(cFormat) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;

        let oNumFormat = AscCommon.oNumFormatCache.get(cFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm._apiForm.value;
        let oFormatParser = AscCommon.g_oFormatParser;

        function getShortPattern(aRawFormat) {
            let dayDone     = false;
            let monthDone   = false;
            let yearDone    = false;
            
            let sPattern = "";

            let numFormat_Year = 12;
            let numFormat_Month = 13;
            let numFormat_Day = 16;

            for (let obj of aRawFormat) {
                switch (obj.type) {
                    case numFormat_Day:
                        if (dayDone == false) {
                            sPattern += 1;
                            dayDone = true;
                        }
                        break;
                    case numFormat_Month:
                        if (monthDone == false) {
                            sPattern += 3;
                            monthDone = true;
                        }
                        break;
                    case numFormat_Year:
                        if (yearDone == false) {
                            if (obj.val > 2)
                                sPattern += 5;
                            else
                                sPattern += 4;
                            yearDone = true;
                        }
                        break;
                            
                }
            }
            return sPattern;
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);
        if (null == oNumFormat.oTextFormat.ShortDatePattern) {
            oNumFormat.oTextFormat.ShortDatePattern = getShortPattern(oNumFormat.oTextFormat.aRawFormat);
            oNumFormat.oTextFormat._prepareFormatDatePDF();
        }
        oCultureInfo.ShortDatePattern = oNumFormat.oTextFormat.ShortDatePattern;

        if (oCultureInfo.ShortDatePattern.indexOf("1") == -1)
            oNumFormat.oTextFormat.bDay = false;

        oCultureInfo.AbbreviatedMonthNames.length = 12;
        oCultureInfo.MonthNames.length = 12;

        let oResParsed = oFormatParser.parseDatePDF(sCurValue, oCultureInfo, oNumFormat);
        
        if (sCurValue == "")
            oTargetRun.ClearContent();
        if (!oResParsed) {
            return false;
        }

        oNumFormat.oTextFormat.formatType = AscCommon.NumFormatType.PDFFormDate;
        let sRes = oNumFormat.oTextFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, oCultureInfo)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Keystroke(cFormat) {
        return true;
    }
    let AFDate_FormatEx = AFDate_Format;
    let AFDate_KeystrokeEx = AFDate_Keystroke;

    /**
	 * Convert field value to specific time format.
	 * @memberof CTextField
     * @param {number} ptf - time format
     *  0 = 24HR_MM [ 14:30 ]
     *  1 = 12HR_MM [ 2:30 PM ]
     *  2 = 24HR_MM_SS [ 14:30:15 ]
     *  3 = 12HR_MM_SS [ 2:30:15 PM ]
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Format(ptf) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        if (!oCurForm)
            return;

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);

        let sFormat;
        let fIsValidTime = null;
        switch (ptf) {
            case 0:
                sFormat = "HH:MM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*$/;
                    return pattern.test(time);
                }
                break;
            case 1:
                sFormat = "h:MM AM/PM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
            case 2:
                sFormat = "HH:MM:ss";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*$/;
                    return pattern.test(time);
                }
                break;
            case 3:
                sFormat = "h:MM:ss AM/PM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
        }

        let oNumFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);
        let sCurValue = oCurForm._apiForm.value;
        
        if (sCurValue == "")
            oTargetRun.ClearContent();
        else if (fIsValidTime(sCurValue) == false)
            return false;

        let oFormatParser = AscCommon.g_oFormatParser;
        let oResParsed = oFormatParser.parseDatePDF(sCurValue, AscCommon.g_aCultureInfos[9]);

        if (!oResParsed) {
            oTargetRun.ClearContent();
            return false;
        }
        
        oNumFormat.oTextFormat.formatType = AscCommon.NumFormatType.PDFFormDate;
        let sRes = oNumFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, true, undefined, true)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Keystroke(cFormat) {
        return true;
    }

    let AFTime_FormatEx = AFDate_FormatEx;
    let AFTime_KeystrokeEx = AFTime_Keystroke;

    /**
	 * Convert field value to specific special format.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Format(psf) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        if (!oCurForm)
            return;

        let sFormValue = oCurForm._apiForm.value;
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        function isValidZipCode(zipCode) {
            let regex = /^\d{5}$/;
            return regex.test(zipCode);
        }
        function isValidZipCode4(zip) {
            let regex = /^\d{5}[-\s.]?(\d{4})?$/;
            return regex.test(zip);
        }
        function isValidPhoneNumber(number) {
            let regex = /^\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/;
            return regex.test(number);
        }
        function isValidSSN(ssn) {
            let regex = /^\d{3}[-\s.]?\d{2}[-\s.]?\d{4}$/;
            return regex.test(ssn);
        }

        if (sFormValue == "")
            oTargetRun.ClearContent();
            
        switch (psf) {
            case 0:
                if (isValidZipCode(sFormValue) == false)
                    return false;
                break;
            case 1:
                if (isValidZipCode4(sFormValue) == false)
                    return false;
                break;
            case 2:
                if (isValidPhoneNumber(sFormValue) == false)
                    return false;
                break;
            case 3:
                if (isValidSSN(sFormValue) == false)
                    return false;
                break;
        }

        sFormValue = sFormValue.replace(/\D/g, ""); // delete all except no digit chars
        let sFormatValue = "";

        oTargetRun.ClearContent();

        switch (psf) {
            case 0:
                sFormatValue = sFormValue.substring(0, 5);
                break;
            case 1:
                sFormatValue = sFormValue.substring(0, 9);
                if (sFormatValue[4])
                    sFormatValue = sFormValue.substring(0, 5) + "-" + sFormValue.substring(5);
                break;
            case 2: 
                let x = sFormValue.substring(0, 10);
                sFormatValue = x.length > 6
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3, 6) + "-" + x.substring(6, 10)
                    : x.length > 3
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3)
                    : x;
                break;
            case 3:
                let y = sFormValue.substring(0, 9);
                sFormatValue = y.length > 5
                ? y.substring(0, 3) + "-" + y.substring(3, 5) + "-" + y.substring(5, 9)
                : y.length > 2
                ? y.substring(0, 3) + "-" + x.substring(3)
                : x;
                break;
        }

        oTargetRun.AddText(sFormatValue);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Keystroke(psf) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidZipCode(zipCode) {
            let regex = /^\d{0,5}$/;
            return regex.test(zipCode);
        }
        function isValidZipCode4(zip) {
            let regex = /^\d{0,5}[-\s.]?(\d{0,4})?$/;
            return regex.test(zip);
        }
        function isValidPhoneNumber(number) {
            let regex = /^\(?\d{0,3}?\)?[\s.-]?\d{0,3}?[\s.-]?\d{0,4}?$/;
            return regex.test(number);
        }
        function isValidSSN(ssn) {
            let regex = /^\d{0,3}?[-\s.]?\d{0,2}?[-\s.]?\d{0,4}$/;
            return regex.test(ssn);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});
        
        let canAdd;
        switch (psf) {
            case 0:
                canAdd = isValidZipCode(sResultText);
                break;
            case 1:
                canAdd = isValidZipCode4(sResultText);
                break;
            case 2:
                canAdd = isValidPhoneNumber(sResultText);
                break;
            case 3:
                canAdd = isValidSSN(sResultText);
                break;
        }

        return canAdd;
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} mask - the special mask
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_KeystrokeEx(mask) {
        let oDoc = private_getViewer().doc;
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});
        
        if (typeof(mask) == "string" && mask != "") {
            let oTextFormat = new AscWord.CTextFormFormat();
            let arrBuffer = oTextFormat.GetBuffer(sResultText);

            let oFormMask = new AscWord.CTextFormMask();
            oFormMask.Set(mask);
            return oFormMask.Check(arrBuffer);
        }

        return false;
    }


    
    // private methods

    function private_GetColor(r, g, b, Auto)
	{
		return new AscCommonWord.CDocumentColor(r, g, b, Auto ? Auto : false);
	}
    
    function private_doKeystrokeAction(oField, aChars) {
        let isValid = true;
        if (oField._actions.Keystroke) {
            oField._doc.enteredFormChars = aChars;
            oField._doc.activeForm = oField;
            isValid = eval(oField._actions.Keystroke.script);
        }

        return isValid;
    }

    function private_GetFieldAlign(sJc)
	{
		if ("left" === sJc)
			return align_Left;
		else if ("right" === sJc)
			return align_Right;
		else if ("center" === sJc)
			return align_Center;

		return undefined;
	}

    function private_getPageCoordsMM(x, y, nPage) {
        // конвертация из глобальных x, y к mm кординатам самой страницы
        let oViewer = private_getViewer();
        var pageObject = oViewer.getPageByCoords(x - oViewer.x, y - oViewer.y);

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let X = (pageObject.x) * g_dKoef_pix_to_mm * nScaleY;
        let Y = (pageObject.y) * g_dKoef_pix_to_mm * nScaleX;

        return {X, Y};
    }
    function private_getGlobalCoordsByPageCoords(x, y, nPage) {
        let oViewer = private_getViewer();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let X = x * g_dKoef_mm_to_pix / nScaleX;
        let Y = y * g_dKoef_mm_to_pix / nScaleY;

        let pageCoords = oViewer.pageDetector.pages[nPage];

        X = (X * pageCoords.w / oViewer.file.pages[nPage].W + pageCoords.x) / AscCommon.AscBrowser.retinaPixelRatio;
        Y = (Y * pageCoords.h / oViewer.file.pages[nPage].H + pageCoords.y) / AscCommon.AscBrowser.retinaPixelRatio;

        return {X, Y};
    }


    function private_getViewer() {
        return editor.getDocumentRenderer();
    }
    function CreateNewHistoryPointForField(oField) {
        if (AscCommon.History.IsOn() == false)
            AscCommon.History.TurnOn();
        AscCommon.History.Create_NewPoint();
        AscCommon.History.SetAdditionalFormFilling(oField);
    }

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    CComboBoxField.prototype.Remove                 = CTextField.prototype.Remove;
    CComboBoxField.prototype.MoveCursorLeft         = CTextField.prototype.MoveCursorLeft;
    CComboBoxField.prototype.MoveCursorRight        = CTextField.prototype.MoveCursorRight;
    CComboBoxField.prototype.SelectionSetStart      = CTextField.prototype.SelectionSetStart;
    CComboBoxField.prototype.SelectionSetEnd        = CTextField.prototype.SelectionSetEnd;

    CComboBoxField.prototype.SetAlign = CTextField.prototype.SetAlign;
    CComboBoxField.prototype.SetDoNotSpellCheck = CTextField.prototype.SetDoNotSpellCheck;

    CComboBoxField.prototype.RemoveNotAppliedChangesPoints = CTextField.prototype.RemoveNotAppliedChangesPoints;
    CTextField.prototype.UpdateScroll               = CListBoxField.prototype.UpdateScroll;
    CTextField.prototype.ScrollVertical             = CListBoxField.prototype.ScrollVertical;
    
    if (!window["AscPDFEditor"])
	    window["AscPDFEditor"] = {};
        
	window["AscPDFEditor"].FIELD_TYPE           = FIELD_TYPE;
	window["AscPDFEditor"].CPushButtonField     = CPushButtonField;
	window["AscPDFEditor"].CBaseField           = CBaseField;
	window["AscPDFEditor"].CTextField           = CTextField;
	window["AscPDFEditor"].CCheckBoxField       = CCheckBoxField;
	window["AscPDFEditor"].CComboBoxField       = CComboBoxField;
	window["AscPDFEditor"].CListBoxField        = CListBoxField;
	window["AscPDFEditor"].CRadioButtonField    = CRadioButtonField;
	window["AscPDFEditor"].CSignatureField      = CSignatureField;
	window["AscPDFEditor"].private_getGlobalCoordsByPageCoords      = private_getGlobalCoordsByPageCoords;
})();
