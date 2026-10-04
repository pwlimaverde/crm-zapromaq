<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/cursorlocationenum?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: CursorLocationEnum - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/cursorlocationenum?view=sql-server-ver15
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
description: CursorLocationEnum
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 6edb0009-7a69-bb59-a00f-297ecc548feb
document\_version\_independent\_id: 24f9562b-33c8-08d6-10c0-596461544a31
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/cursorlocationenum.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/cursorlocationenum.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 155
asset\_id: ado/reference/ado-api/cursorlocationenum
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/cursorlocationenum.md
platformId: 08a65a78-17d3-3eb6-09f7-cf87913f2587
---
# CursorLocationEnum - ActiveX Data Objects (ADO) | Microsoft Learn
Specifies the location of the cursor service.
| Constant | Value | Description |
| --- | --- | --- |
| \*\*adUseClient\*\* | 3 | Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym \*\*adUseClientBatch\*\* is also supported. |
| \*\*adUseNone\*\* | 1 | Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.) |
| \*\*adUseServer\*\* | 2 | Default. Uses cursors supplied by the data provider or driver. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the [The Microsoft Cursor Service for OLE DB](../../guide/data/the-microsoft-cursor-service-for-ole-db), such as disassociated[Recordset](recordset-object-ado) objects, cannot be simulated with server-side cursors and these features will be unavailable with this setting. |
## ADO/WFC Equivalent
Package: \*\*com.ms.wfc.data\*\*
| Constant |
| --- |
| AdoEnums.CursorLocation.CLIENT |
| AdoEnums.CursorLocation.NONE |
| AdoEnums.CursorLocation.SERVER |
## Applies To
[CursorLocation Property (ADO)](cursorlocation-property-ado)
