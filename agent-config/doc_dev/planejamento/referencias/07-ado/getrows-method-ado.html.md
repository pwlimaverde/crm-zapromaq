<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/getrows-method-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver16
versioningType: Ranged
title: GetRows Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/getrows-method-ado?view=sql-server-ver16
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
description: GetRows Method (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: b454fe6e-45d3-2273-cef7-f1fc17aa2bbb
document\_version\_independent\_id: d9bcc5de-f4c7-a1c0-f6ac-31f2d726fcc2
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/getrows-method-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/getrows-method-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver16
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 292
asset\_id: ado/reference/ado-api/getrows-method-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/getrows-method-ado.md
platformId: 0f0b2308-5404-d209-67fa-877fc2eda971
---
# GetRows Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Retrieves multiple records of a [Recordset](recordset-object-ado) object into an array.
## Syntax
```
array = recordset.GetRows(Rows, Start, Fields )
```
## Return Value
Returns a \*\*Variant\*\* whose value is a two-dimensional array.
#### Parameters
\*Rows\* Optional. A [GetRowsOptionEnum](getrowsoptionenum) value that indicates the number of records to retrieve. The default is \*\*adGetRowsRest\*\*.
\*Start\* Optional. A \*\*String\*\* value or \*\*Variant\*\* that evaluates to the bookmark for the record from which the \*\*GetRows\*\* operation should begin. You can also use a [BookmarkEnum](bookmarkenum) value.
\*Fields\* Optional. A \*\*Variant\*\* that represents a single field name or ordinal position, or an array of field names or ordinal position numbers. ADO returns only the data in these fields.
## Remarks
Use the \*\*GetRows\*\* method to copy records from a \*\*Recordset\*\* into a two-dimensional array. The first subscript identifies the field and the second identifies the record number. The \*array\* variable is automatically dimensioned to the correct size when the \*\*GetRows\*\* method returns the data.
If you do not specify a value for the \*Rows\* argument, the \*\*GetRows\*\* method automatically retrieves all the records in the \*\*Recordset\*\* object. If you request more records than are available, \*\*GetRows\*\* returns only the number of available records.
If the \*\*Recordset\*\* object supports bookmarks, you can specify at which record the \*\*GetRows\*\* method should begin retrieving data by passing the value of that record's [Bookmark](bookmark-property-ado) property in the \*Start\* argument.
If you want to restrict the fields that the \*\*GetRows\*\* call returns, you can pass either a single field name/number or an array of field names/numbers in the \*Fields\* argument.
After you call \*\*GetRows\*\*, the next unread record becomes the current record, or the [EOF](bof-eof-properties-ado) property is set to \*\*True\*\* if there are no more records.
## Applies To
[Recordset Object (ADO)](recordset-object-ado)
---
## Other Supported Versions
- [sql-server-ver16](https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/getrows-method-ado?view=sql-server-ver16&accept=text/markdown)
