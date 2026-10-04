<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/datatypeenum?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: DataTypeEnum - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/datatypeenum?view=sql-server-ver15
config\_moniker\_range: sql-server-ver15
uhfHeaderId: MSDocsHeader-Archive
toc\_preview: true
feedback\_system: None
is\_archived: true
is\_retired: true
ROBOTS: NOINDEX,NOFOLLOW
docs\_archive: tool
feedback\_help\_link\_url: https://learn.microsoft.com/answers/tags/191/sql-server
feedback\_help\_link\_type: get-help-at-qna
recommendations: true
manager: jroth
ms.author: jroth
breadcrumb\_path: ../../../connect/toc.json
tilde\_path: sql-server-ver15
description: DataTypeEnum
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: fcf8b3d8-eae4-1a18-b6db-bf8dc0791e71
document\_version\_independent\_id: 058e5876-101c-174c-1421-1f6741e1025c
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/datatypeenum.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/datatypeenum.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 563
asset\_id: ado/reference/ado-api/datatypeenum
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/datatypeenum.md
platformId: 99ce874c-7b17-8225-dc1c-b016a6973c94
---
# DataTypeEnum - ActiveX Data Objects (ADO) | Microsoft Learn
Specifies the data type of a [Field](field-object), [Parameter](parameter-object), or [Property](property-object-ado). The corresponding OLE DB type indicator is shown in parentheses in the description column of the following table.
| Constant | Value | Description |
| --- | --- | --- |
| \*\*AdArray\*\* | 0x2000 | A flag value, always combined with another data type constant, that indicates an array of the other data type. Does not apply to ADOX. |
| \*\*adBigInt\*\* | 20 | Indicates an eight-byte signed integer (DBTYPE\\_I8). |
| \*\*adBinary\*\* | 128 | Indicates a binary value (DBTYPE\\_BYTES). |
| \*\*adBoolean\*\* | 11 | Indicates a \*\*Boolean\*\* value (DBTYPE\\_BOOL). |
| \*\*adBSTR\*\* | 8 | Indicates a null-terminated character string (Unicode) (DBTYPE\\_BSTR). |
| \*\*adChapter\*\* | 136 | Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE\\_HCHAPTER). |
| \*\*adChar\*\* | 129 | Indicates a string value (DBTYPE\\_STR). |
| \*\*adCurrency\*\* | 6 | Indicates a currency value (DBTYPE\\_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000. |
| \*\*adDate\*\* | 7 | Indicates a date value (DBTYPE\\_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. |
| \*\*adDBDate\*\* | 133 | Indicates a date value (yyyymmdd) (DBTYPE\\_DBDATE). |
| \*\*adDBTime\*\* | 134 | Indicates a time value (hhmmss) (DBTYPE\\_DBTIME). |
| \*\*adDBTimeStamp\*\* | 135 | Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE\\_DBTIMESTAMP). |
| \*\*adDecimal\*\* | 14 | Indicates an exact numeric value with a fixed precision and scale (DBTYPE\\_DECIMAL). |
| \*\*adDouble\*\* | 5 | Indicates a double-precision floating-point value (DBTYPE\\_R8). |
| \*\*adEmpty\*\* | 0 | Specifies no value (DBTYPE\\_EMPTY). |
| \*\*adError\*\* | 10 | Indicates a 32-bit error code (DBTYPE\\_ERROR). |
| \*\*adFileTime\*\* | 64 | Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE\\_FILETIME). |
| \*\*adGUID\*\* | 72 | Indicates a globally unique identifier (GUID) (DBTYPE\\_GUID). |
| \*\*adIDispatch\*\* | 9 | Indicates a pointer to an \*\*IDispatch\*\* interface on a COM object (DBTYPE\\_IDISPATCH).\*\*Note\*\* This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| \*\*adInteger\*\* | 3 | Indicates a four-byte signed integer (DBTYPE\\_I4). |
| \*\*adIUnknown\*\* | 13 | Indicates a pointer to an \*\*IUnknown\*\* interface on a COM object (DBTYPE\\_IUNKNOWN).\*\*Note\*\* This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| \*\*adLongVarBinary\*\* | 205 | Indicates a long binary value. |
| \*\*adLongVarChar\*\* | 201 | Indicates a long string value. |
| \*\*adLongVarWChar\*\* | 203 | Indicates a long null-terminated Unicode string value. |
| \*\*adNumeric\*\* | 131 | Indicates an exact numeric value with a fixed precision and scale (DBTYPE\\_NUMERIC). |
| \*\*adPropVariant\*\* | 138 | Indicates an Automation PROPVARIANT (DBTYPE\\_PROP\\_VARIANT). |
| \*\*adSingle\*\* | 4 | Indicates a single-precision floating-point value (DBTYPE\\_R4). |
| \*\*adSmallInt\*\* | 2 | Indicates a two-byte signed integer (DBTYPE\\_I2). |
| \*\*adTinyInt\*\* | 16 | Indicates a one-byte signed integer (DBTYPE\\_I1). |
| \*\*adUnsignedBigInt\*\* | 21 | Indicates an eight-byte unsigned integer (DBTYPE\\_UI8). |
| \*\*adUnsignedInt\*\* | 19 | Indicates a four-byte unsigned integer (DBTYPE\\_UI4). |
| \*\*adUnsignedSmallInt\*\* | 18 | Indicates a two-byte unsigned integer (DBTYPE\\_UI2). |
| \*\*adUnsignedTinyInt\*\* | 17 | Indicates a one-byte unsigned integer (DBTYPE\\_UI1). |
| \*\*adUserDefined\*\* | 132 | Indicates a user-defined variable (DBTYPE\\_UDT). |
| \*\*adVarBinary\*\* | 204 | Indicates a binary value. |
| \*\*adVarChar\*\* | 200 | Indicates a string value. |
| \*\*adVariant\*\* | 12 | Indicates an Automation \*\*Variant\*\* (DBTYPE\\_VARIANT).\*\*Note\*\* This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| \*\*adVarNumeric\*\* | 139 | Indicates a numeric value. |
| \*\*adVarWChar\*\* | 202 | Indicates a null-terminated Unicode character string. |
| \*\*adWChar\*\* | 130 | Indicates a null-terminated Unicode character string (DBTYPE\\_WSTR). |
## ADO/WFC Equivalent
Package: \*\*com.ms.wfc.data\*\*
| Constant |
| --- |
| AdoEnums.DataType.ARRAY |
| AdoEnums.DataType.BIGINT |
| AdoEnums.DataType.BINARY |
| AdoEnums.DataType.BOOLEAN |
| AdoEnums.DataType.BSTR |
| AdoEnums.DataType.CHAPTER |
| AdoEnums.DataType.CHAR |
| AdoEnums.DataType.CURRENCY |
| AdoEnums.DataType.DATE |
| AdoEnums.DataType.DBDATE |
| AdoEnums.DataType.DBTIME |
| AdoEnums.DataType.DBTIMESTAMP |
| AdoEnums.DataType.DECIMAL |
| AdoEnums.DataType.DOUBLE |
| AdoEnums.DataType.EMPTY |
| AdoEnums.DataType.ERROR |
| AdoEnums.DataType.FILETIME |
| AdoEnums.DataType.GUID |
| AdoEnums.DataType.IDISPATCH |
| AdoEnums.DataType.INTEGER |
| AdoEnums.DataType.IUNKNOWN |
| AdoEnums.DataType.LONGVARBINARY |
| AdoEnums.DataType.LONGVARCHAR |
| AdoEnums.DataType.LONGVARWCHAR |
| AdoEnums.DataType.NUMERIC |
| AdoEnums.DataType.PROPVARIANT |
| AdoEnums.DataType.SINGLE |
| AdoEnums.DataType.SMALLINT |
| AdoEnums.DataType.TINYINT |
| AdoEnums.DataType.UNSIGNEDBIGINT |
| AdoEnums.DataType.UNSIGNEDINT |
| AdoEnums.DataType.UNSIGNEDSMALLINT |
| AdoEnums.DataType.UNSIGNEDTINYINT |
| AdoEnums.DataType.USERDEFINED |
| AdoEnums.DataType.VARBINARY |
| AdoEnums.DataType.VARCHAR |
| AdoEnums.DataType.VARIANT |
| AdoEnums.DataType.VARNUMERIC |
| AdoEnums.DataType.VARWCHAR |
| AdoEnums.DataType.WCHAR |
## Applies To
[Append Method (ADO)](append-method-ado)[CreateParameter Method (ADO)](createparameter-method-ado)
[CreateRecordset Method (RDS)](../rds-api/createrecordset-method-rds)[Type Property (ADO)](type-property-ado)
