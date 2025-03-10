/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

"use strict";

(function(window, undefined){
    var oMonths = {};
    oMonths[0] = "January";
    oMonths[1] = "February";
    oMonths[2] = "March";
    oMonths[3] = "April";
    oMonths[4] = "May";
    oMonths[5] = "June";
    oMonths[6] = "July";
    oMonths[7] = "August";
    oMonths[8] = "September";
    oMonths[9] = "October";
    oMonths[10] = "November";
    oMonths[11] = "December";

    var oDays = {};
    oDays[0] = "Sunday";
    oDays[1] = "Monday";
    oDays[2] = "Tuesday";
    oDays[3] = "Wednesday";
    oDays[4] = "Thursday";
    oDays[5] = "Friday";
    oDays[6] = "Saturday";



    var oDateTimeFormats = {};
    oDateTimeFormats["datetime1"] = "MM/DD/YYYY";
    oDateTimeFormats["datetimeFigureOut"] = oDateTimeFormats["datetime1"];
    oDateTimeFormats["datetime2"] = "dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
    oDateTimeFormats["datetime3"] = "DD\\ MMMM\\ YYYY";
    oDateTimeFormats["datetime4"] = "MMMM\\ DD\\,\\ YYYY";
    oDateTimeFormats["datetime5"] = "DD-MMM-YY";
    oDateTimeFormats["datetime6"] = "MMMM\\ YY";
    oDateTimeFormats["datetime7"] = "MMM-YY";
    oDateTimeFormats["datetime8"] = "MM/DD/YYYY\\ hh:mm\\ AM/PM";
    oDateTimeFormats["datetime9"] = "MM/DD/YYYY\\ hh:mm:ss\\ AM/PM";
    oDateTimeFormats["datetime10"] = "hh:mm";
    oDateTimeFormats["datetime11"] = "hh:mm:ss";
    oDateTimeFormats["datetime12"] = "hh:mm\\ AM/PM";
    oDateTimeFormats["datetime13"] = "hh:mm:ss:\\ AM/PM";

    function CPresentationField(Paragraph)
    {
        ParaRun.call(this, Paragraph, false);
        this.Guid = null;
        this.FieldType = null;
        this.PPr = null;

        this.Slide = null;
        this.SlideNum = null;
        this.CanAddToContent = false;
    }
    CPresentationField.prototype = Object.create(ParaRun.prototype);
    CPresentationField.prototype.constructor = CPresentationField;


    CPresentationField.prototype.Copy = function(Selected, oPr)
    {
        if(oPr && oPr.Paragraph && oPr.Paragraph.bFromDocument)
        {
            return ParaRun.prototype.Copy.call(this, Selected, oPr);
        }
        var Field = new CPresentationField(this.Paragraph);
        Field.Set_Pr( this.Pr.Copy() );
        Field.SetGuid(AscCommon.CreateGUID());
        Field.SetFieldType( this.FieldType );
        if(this.PPr)
        {
            Field.SetPPr(this.PPr.Copy());
        }
        return Field;
    };

    CPresentationField.prototype.Copy2 = function()
    {
        this.Copy();
    };

    CPresentationField.prototype.SetGuid = function(sGuid)
    {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PresentationField_Guid, this.Guid, sGuid));
        this.Guid = sGuid;
    };
    CPresentationField.prototype.SetFieldType = function(Type)
    {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PresentationField_FieldType, this.FieldType, Type));
        this.FieldType = Type;
    };
    CPresentationField.prototype.SetPPr = function(Pr)
    {
        History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_PresentationField_PPr, this.PPr, Pr));
        this.PPr = Pr;
    };

    CPresentationField.prototype.Add_ToContent = function(Pos, Item, UpdatePosition)
    {
        if(AscCommon.History.Is_On() && !this.CanAddToContent)
        {
            return;
        }
        ParaRun.prototype.Add_ToContent.call(this, Pos, Item, UpdatePosition);
    };
    CPresentationField.prototype.Remove_FromContent = function(Pos, Count, UpdatePosition)
    {
        if(AscCommon.History.Is_On())
        {
            return;
        }
        ParaRun.prototype.Remove_FromContent.call(this, Pos, Count, UpdatePosition);
    };
    CPresentationField.prototype.Is_Empty = function()
    {
        return false;
    };

    CPresentationField.prototype.private_CalculateContent = function()
    {
        AscFormat.ExecuteNoHistory(function(){
            var sStr = this.private_GetString();
            if(typeof sStr === 'string')
            {
                const bSelectionUse = this.IsSelectionUse();
                const oSelection = this.State.Selection;
                const nDirection = oSelection.EndPos - oSelection.StartPos;
                this.Content.length = 0;
                this.AddText(sStr, -1);
                if(bSelectionUse)
                {
                    this.SelectAll(nDirection);
                }
            }
        }, this, []);
    };
    CPresentationField.prototype.GetFieldType = function(){
        if(typeof this.FieldType === 'string') {
            return this.FieldType.toLowerCase();
        }
        return "";
    };
    CPresentationField.prototype.private_GetString = function()
    {
        var sStr = null;
        var oStylesObject;
        var oCultureInfo = AscCommon.g_aCultureInfos[this.Get_CompiledPr().Lang.Val];
        if(!oCultureInfo)
        {
            oCultureInfo = AscCommon.g_aCultureInfos[1033];
        }
        var oDateTime, oFormat;
        if(typeof this.FieldType === 'string')
        {
            var sFieldType = this.FieldType.toLowerCase();
            sStr = null;
            if("slidenum" === sFieldType)
            {
                if(this.Paragraph && this.Paragraph.Parent)
                {
                    oStylesObject = this.Paragraph.Parent.Get_Styles(0);
                    var nFirstSlideNum = 1;
                    if(oStylesObject.presentation)
                    {
                        nFirstSlideNum = oStylesObject.presentation.getFirstSlideNumber();
                    }
                    if(oStylesObject.slide)
                    {
                        this.Slide = oStylesObject.slide;
                        if(AscFormat.isRealNumber(this.Slide.num))
                        {
                            this.SlideNum = this.Slide.num;
                            sStr = '' + (this.Slide.num + nFirstSlideNum);
                        }
                    }
                    else if(oStylesObject.notes)
                    {
                        if(oStylesObject.notes.slide)
                        {
                            this.Slide = oStylesObject.notes.slide;
                            if(AscFormat.isRealNumber(this.Slide.num))
                            {
                                this.SlideNum = this.Slide.num;
                                sStr = '' + (this.Slide.num + nFirstSlideNum);
                            }
                        }
                    }
                    else if(oStylesObject.layout || oStylesObject.master)
                    {
                        let oParent = (oStylesObject.layout || oStylesObject.master);
                        if(AscFormat.isRealNumber(oParent.lastRecalcSlideIndex) && oParent.lastRecalcSlideIndex > -1)
                        {
                            this.SlideNum = oParent.lastRecalcSlideIndex;
                            sStr = '' + (this.SlideNum + nFirstSlideNum);
                        }
                        else
                        {
                            sStr = '<#>';
                        }
                    }
                }
            }
            else if("value" === sFieldType)
            {
                if(this.Paragraph && this.Paragraph.Parent)
                {
                    oStylesObject = this.Paragraph.Parent.Get_Styles();
                    if(oStylesObject.shape && oStylesObject.shape.getValueString && oStylesObject.shape.getValueString())
                    {
                        sStr = oStylesObject.shape.getValueString();
                    }
                }
            }
            else if("categoryname" === sFieldType)
            {
                if(this.Paragraph && this.Paragraph.Parent)
                {
                    oStylesObject = this.Paragraph.Parent.Get_Styles();
                    if(oStylesObject.shape && oStylesObject.shape.getCategoryName && oStylesObject.shape.getCategoryName())
                    {
                        sStr = oStylesObject.shape.getCategoryName();
                    }
                }
            }
            else if("percentage" === sFieldType)
            {
                if(this.Paragraph && this.Paragraph.Parent)
                {
                    oStylesObject = this.Paragraph.Parent.Get_Styles();
                    if(oStylesObject.shape && oStylesObject.shape.getPercentageString())
                    {
                        sStr = oStylesObject.shape.getPercentageString();
                    }
                }
            }
            else if(sFieldType.indexOf("datetime") === 0)
            {
                oFormat = this.private_GetDateTimeFormat(sFieldType);
                if(oFormat)
                {
                    oDateTime = new Asc.cDate();
                    sStr =  oFormat.formatToWord(oDateTime.getExcelDate(true) + (oDateTime.getHours() * 60 * 60 + oDateTime.getMinutes() * 60 + oDateTime.getSeconds()) / AscCommonExcel.c_sPerDay, 15, oCultureInfo);
                }
                else
                {
                    sStr = sFieldType.toUpperCase();
                }
            }
            else
            {
                sStr = sFieldType.toUpperCase();
            }
        }
        return sStr;
    };

    CPresentationField.prototype.private_GetDateTimeFormat = function(sFieldType)
    {
        let sResultFiledType = sFieldType;
        let oFormat = null;
        const nLang = this.Get_CompiledPr().Lang.Val;
        let sFormat = oDateTimeFormats[sResultFiledType] || AscCommonWord.oDefaultDateTimeFormat[nLang];
        if(!sFormat)
        {
            sResultFiledType = "datetime1";
            sFormat = oDateTimeFormats[sResultFiledType]
        }
        if(sFormat)
        {
            let aFormat = Asc.c_oAscDateTimeFormat[nLang];
            if(!Array.isArray(aFormat))
            {
                aFormat = Asc.c_oAscDateTimeFormat[lcid_enUS];
            }
            if(Array.isArray(aFormat))
            {
                let nIdx = 0;
                //match field type to index in Asc.c_oAscDateTimeFormat[nLang]
                switch (sResultFiledType)
                {
                    case "datetimeFigureOut": nIdx = 0; break;//"MM/DD/YYYY";
                    case "datetime1": nIdx = 0; break;//"MM/DD/YYYY";
                    case "datetime2": nIdx = 1; break;//"dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
                    case "datetime3": nIdx = 8; break;//"DD\\ MMMM\\ YYYY";
                    case "datetime4": nIdx = 2; break;//"MMMM\\ DD\\,\\ YYYY";
                    case "datetime5": nIdx = 5; break;//"DD-MMM-YY";
                    case "datetime6": nIdx = 9; break;//"MMMM\\ YY";
                    case "datetime7": nIdx = 10; break;//"MMM-YY";
                    case "datetime8": nIdx = 11; break;//"MM/DD/YYYY\\ hh:mm\\ AM/PM";
                    case "datetime9": nIdx = 12; break;//"MM/DD/YYYY\\ hh:mm:ss\\ AM/PM";
                    case "datetime10": nIdx = 15; break;//"hh:mm";
                    case "datetime11": nIdx = 16; break;//"hh:mm:ss";
                    case "datetime12": nIdx = 13; break;//"hh:mm\\ AM/PM";
                    case "datetime13": nIdx = 14; break;//"hh:mm:ss:\\ AM/PM";
                }
                if(aFormat[nIdx])
                {
                    sFormat = aFormat[nIdx]
                }
            }
            oFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.WordFieldDate);
        }
        return oFormat;
    };

    CPresentationField.prototype.Recalculate_MeasureContent = function()
    {
        if (!this.RecalcInfo.IsMeasureNeed())
            return;
        this.private_CalculateContent();
        ParaRun.prototype.Recalculate_MeasureContent.call(this);
    };

    CPresentationField.prototype.Recalculate_MeasureContent = function()
    {
        if (!this.RecalcInfo.IsMeasureNeed())
            return;
        this.private_CalculateContent();
        ParaRun.prototype.Recalculate_MeasureContent.call(this);
    };

    CPresentationField.prototype.Write_ToBinary2 = function(Writer)
    {
        var StartPos = Writer.GetCurPosition();
        ParaRun.prototype.Write_ToBinary2.call(this, Writer);
        var EndPos = Writer.GetCurPosition();
        Writer.Seek(StartPos);
        Writer.WriteLong( AscDFH.historyitem_type_PresentationField);
        Writer.Seek(EndPos);
    };
    CPresentationField.prototype.GetSelectedElementsInfo = function(oInfo)
	{
		oInfo.SetPresentationField(this);
		ParaRun.prototype.GetSelectedElementsInfo.apply(this, arguments);
	};
	CPresentationField.prototype.Set_SelectionContentPos = function(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag)
	{
		if (this.Paragraph && this.Paragraph.GetSelectDirection() > 0)
			this.SelectAll(1);
		else
			this.SelectAll(-1);
	};
	CPresentationField.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
	{
		if (false === UseContentPos && this.Content.length > 0)
		{
			SearchPos.Found = true;
			SearchPos.Pos.Update(0, Depth);
			return true;
		}

		return false;
	};
	CPresentationField.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
	{
		if (false === UseContentPos && this.Content.length > 0)
		{
			SearchPos.Found = true;
			SearchPos.Pos.Update(this.Content.length, Depth);
			return true;
		}

		return false;
	};
	CPresentationField.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
	{
	};
	CPresentationField.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
	{
	};
	CPresentationField.prototype.IsSolid = function()
	{
		return true;
	};
	CPresentationField.prototype.IsStopCursorOnEntryExit = function()
	{
		return true;
	};
	CPresentationField.prototype.Cursor_Is_NeededCorrectPos = function()
	{
		return false;
	};

    var drawingsChangesMap = window['AscDFH'].drawingsChangesMap;
    drawingsChangesMap[AscDFH.historyitem_PresentationField_FieldType] = function(oClass, value){oClass.FieldType = value;};
    drawingsChangesMap[AscDFH.historyitem_PresentationField_Guid] = function(oClass, value){oClass.Guid = value;};
    drawingsChangesMap[AscDFH.historyitem_PresentationField_PPr] = function(oClass, value){oClass.PPr = value;};

    AscDFH.changesFactory[AscDFH.historyitem_PresentationField_FieldType] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_PresentationField_Guid] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_PresentationField_PPr] = window['AscDFH'].CChangesDrawingsObjectNoId;

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CPresentationField = CPresentationField;
window['AscCommonWord'].oDefaultDateTimeFormat = {};
})(window);
