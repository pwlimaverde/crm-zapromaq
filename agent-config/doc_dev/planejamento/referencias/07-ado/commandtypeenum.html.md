<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/commandtypeenum?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: CommandTypeEnum - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/commandtypeenum?view=sql-server-ver15
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
description: CommandTypeEnum
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: c787953c-0934-28c8-2dfe-aafcbcb00056
document\_version\_independent\_id: f51a5b7f-2813-72f6-6ef6-c85e354f08b1
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/commandtypeenum.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/commandtypeenum.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 277
asset\_id: ado/reference/ado-api/commandtypeenum
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/commandtypeenum.md
platformId: cb54e7cb-de1f-2a0c-57c4-93dbf0d9fe78
---
# CommandTypeEnum - ActiveX Data Objects (ADO) | Microsoft Learn
Specifies how a command argument should be interpreted.
It is important to validate user-supplied \*CommandString\* values to avoid giving application users the opportunity to inject potentially dangerous commands for ADO to execute.
| Constant | Value | Description |
| --- | --- | --- |
| \*\*adCmdUnspecified\*\* | -1 | Does not specify the command type argument. |
| \*\*adCmdText\*\* | 1 | Evaluates [CommandText](commandtext-property-ado) as a textual definition of a command or stored procedure call. |
| \*\*adCmdTable\*\* | 2 | Evaluates \*\*CommandText\*\* as a table name whose columns are all returned by an internally generated SQL query. |
| \*\*adCmdStoredProc\*\* | 4 | Evaluates \*\*CommandText\*\* as a stored procedure name. |
| \*\*adCmdUnknown\*\* | 8 | Default. Indicates that the type of command in the \*\*CommandText\*\* property is not known. When the type of command is not known, ADO will make several attempts to interpret the \*\*CommandText\*\*. - \*\*CommandText\*\* is interpreted as a textual definition of a command or stored procedure call. This is the same behavior as \*\*adCmdText\*\*.- \*\*CommandText\*\* is the name of a stored procedure. This is the same behavior as \*\*adCmdStoredProc\*\*.- \*\*CommandText\*\* is interpreted as the name of a table. All columns are returned by an internally generated SQL query. This is the same behavior as \*\*adCmdTable\*\*. |
| \*\*adCmdFile\*\* | 256 | Evaluates \*\*CommandText\*\* as the file name of a persistently stored [Recordset](recordset-object-ado). Used with \*\*Recordset.\*\*[Open](open-method-ado-recordset) or [Requery](requery-method) only. |
| \*\*adCmdTableDirect\*\* | 512 | Evaluates \*\*CommandText\*\* as a table name whose columns are all returned. Used with \*\*Recordset.Open\*\* or \*\*Requery\*\* only. To use the [Seek](seek-method) method, the \*\*Recordset\*\* must be opened with \*\*adCmdTableDirect\*\*. This value cannot be combined with the [ExecuteOptionEnum](executeoptionenum) value \*\*adAsyncExecute\*\*. |
## ADO/WFC Equivalent
Package: \*\*com.ms.wfc.data\*\*
| Constant |
| --- |
| AdoEnums.CommandType.UNSPECIFIED |
| AdoEnums.CommandType.TEXT |
| AdoEnums.CommandType.TABLE |
| AdoEnums.CommandType.STOREDPROC |
| AdoEnums.CommandType.UNKNOWN |
| AdoEnums.CommandType.FILE |
| AdoEnums.CommandType.TABLEDIRECT |
## Applies To
[CommandType Property (ADO)](commandtype-property-ado)[Execute Method (ADO Command)](execute-method-ado-command)
[Execute Method (ADO Connection)](execute-method-ado-connection)[Open Method (ADO Recordset)](open-method-ado-recordset)
[Requery Method](requery-method)
