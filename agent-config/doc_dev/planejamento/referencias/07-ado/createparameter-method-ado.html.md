<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/createparameter-method-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: CreateParameter Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/createparameter-method-ado?view=sql-server-ver15
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
description: CreateParameter Method (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 8784b28d-41f9-de24-550f-36da9f7a1809
document\_version\_independent\_id: d56d62b3-3483-8a62-eb49-58f1c4f0e306
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/createparameter-method-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/createparameter-method-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 255
asset\_id: ado/reference/ado-api/createparameter-method-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/createparameter-method-ado.md
platformId: e352c313-3bc1-2538-bbb8-76ff4eaf353a
---
# CreateParameter Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Creates a new [Parameter](parameter-object) object with the specified properties.
## Syntax
```
Set parameter = command.CreateParameter (Name, Type, Direction, Size, Value)
```
## Return Value
Returns a \*\*Parameter\*\* object.
#### Parameters
\*Name\* Optional. A \*\*String\*\* value that contains the name of the \*\*Parameter\*\* object.
\*Type\* Optional. A [DataTypeEnum](datatypeenum) value that specifies the data type of the \*\*Parameter\*\* object.
\*Direction\* Optional. A [ParameterDirectionEnum](parameterdirectionenum) value that specifies the type of \*\*Parameter\*\* object.
\*Size\* Optional. A \*\*Long\*\* value that specifies the maximum length for the parameter value in characters or bytes.
\*Value\* Optional. A \*\*Variant\*\* that specifies the value for the \*\*Parameter\*\* object.
## Remarks
Use the \*\*CreateParameter\*\* method to create a new \*\*Parameter\*\* object with a specified name, type, direction, size, and value. Any values you pass in the arguments are written to the corresponding \*\*Parameter\*\* properties.
This method does not automatically append the \*\*Parameter\*\* object to the \*\*Parameters\*\* collection of a [Command](command-object-ado) object. This lets you set additional properties whose values ADO will validate when you append the \*\*Parameter\*\* object to the collection.
If you specify a variable-length data type in the \*Type\* argument, you must either pass a \*Size\* argument or set the [Size](size-property-ado-parameter) property of the \*\*Parameter\*\* object before appending it to the \*\*Parameters\*\* collection; otherwise, an error occurs.
If you specify a numeric data type (\*\*adNumeric\*\* or \*\*adDecimal\*\*) in the \*Type\* argument, then you must also set the [NumericScale](numericscale-property-ado) and [Precision](precision-property-ado) properties.
## Applies To
[Command Object (ADO)](command-object-ado)
