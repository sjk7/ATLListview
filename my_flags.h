#pragma once
#ifndef __MY_FLAGS_H_INCLUDED
#define __MY_FLAGS_H_INCLUDED


[uuid(EE9CB4F0-8118-4772-8B91-B9C448DC8A04), v1_enum]
typedef enum ListLabelEditConstants{
	lvwAutomatic,
		lvwManual
} ListLabelEditConstants;

[uuid(985ECCDF-A665-489D-887B-2BF1E5C8156B), v1_enum]
typedef enum BorderStyleConstants {
	ccNone = 0,
	ccFixedSingle = 1
}BorderStyleConstants;

[uuid(550EC01B-FF02-4AC4-A836-4366A35F6248)]
typedef enum ListSortOrderFlags {
	lvwNone = 0,
	lvwAscending = 1,
	lvwDescending = 2,
	lvwNatural = 4,
	lvwNoCase = 8

}ListSortOrderFlags;


[uuid(A81E9973-2781-4CC6-B325-42FB3E2DB3E2), v1_enum]
typedef enum ListColumnAlignmentConstants{
    lvwColumnLeft = 0,
    lvwColumnRight = 1,
    lvwColumnCenter = 2
} ListColumnAlignmentConstants;

[uuid(5407754A-1B28-45A5-8D46-B4B9B5D39B7B), v1_enum]
typedef enum ColumnContentType{
	lvText = 0,
	lvNumber = 1,
	lvDate = 2,
	lvTimeFormat = 4
} ColumnContentType;

[uuid(6CB4DCFB-6598-4726-8A14-C14736529B98), v1_enum]
typedef enum AppearanceConstants {
	ccFlat= 0,
	cc3D = 1
}AppearanceConstants;

[uuid(F690F6F2-9587-4DEE-B8D1-61F5DDA806DE), v1_enum]
typedef enum ShiftConstants{
	vbShiftMask = 1,
	vbCtrlMask = 2,
	vbAltMask = 4
}vbShiftConstants;

[uuid(2454E41D-A2CE-488B-8C1C-FCD7D05AFE4E), v1_enum]
typedef enum ListColumnResizeMode{
    lvwResizeHorizontalProportion = 0,
    lvwResizeKeepWidth = 1
} ListColumnResizeMode;


[uuid(3511FB6C-CB9B-4DE6-A95A-D4554D7885F7), v1_enum]
typedef enum MouseButtonConstants{

	VbLeftButton = 1,
	VbRightButton = 2,
	VbMiddleButton = 4

}vbMouseButtonConstants;


typedef enum {
	fmt_left = 0x0001,
	fmt_right = 0x0002,
	fmt_center = 0x0004,
	fmt_justify_mask = 0x0003,
	fmt_image = 0x0800,
	fmt_bitmap_right = 0x1000,
	fmt_has_images = 0x8000
}FormatFlags;

typedef enum {
	lvcf_fmt = 0x0001,
	lvcf_width = 0x0002,
	lvcf_text = 0x0004,
	lvcf_subitem = 0x0008,
	lvcf_image = 0x0010,
	lvcf_order = 0x0020
}FormatMask;





#endif
