<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameter-object?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: Parameter Object - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameter-object?view=sql-server-ver15
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
description: Parameter Object
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 41a8d50f-0ffc-c071-1a3b-a86d616a480c
document\_version\_independent\_id: 42c2f600-7603-0f89-38cb-a02005f4ad52
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameter-object.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameter-object.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 316
asset\_id: ado/reference/ado-api/parameter-object
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameter-object.md
platformId: 0c0495dd-e499-ea09-94f6-60b3c19ef52e
---
# Parameter Object - ActiveX Data Objects (ADO) | Microsoft Learn
Represents a parameter or argument associated with a [Command](command-object-ado) object based on a parameterized query or stored procedure.
## Remarks
Many providers support parameterized commands. These are commands in which the desired action is defined once, but variables (or parameters) are used to alter some details of the command. For example, a SQL SELECT statement could use a parameter to define the matching criteria of a WHERE clause, and another to define the column name for a SORT BY clause.
\*\*Parameter\*\* objects represent parameters associated with parameterized queries, or the in/out arguments and the return values of stored procedures. Depending on the functionality of the provider, some collections, methods, or properties of a \*\*Parameter\*\* object may not be available.
With the collections, methods, and properties of a \*\*Parameter\*\* object, you can do the following:
- Set or return the name of a parameter with the [Name](name-property-ado) property.
- Set or return the value of a parameter with the [Value](value-property-ado) property. \*\*Value\*\* is the default property of the \*\*Parameter\*\* object.
- Set or return parameter characteristics with the [Attributes](attributes-property-ado), [Direction](direction-property), [Precision](precision-property-ado), [NumericScale](numericscale-property-ado), [Size](size-property-ado-parameter), and [Type](type-property-ado) properties.
- Pass long binary or character data to a parameter with the [AppendChunk](appendchunk-method-ado) method.
- Access provider-specific attributes by using the [Properties](properties-collection-ado) collection.
If you know the names and properties of the parameters associated with the stored procedure or parameterized query you want to call, you can use the [CreateParameter](createparameter-method-ado) method to create \*\*Parameter\*\* objects with the appropriate property settings and use the [Append](append-method-ado) method to add them to the [Parameters](parameters-collection-ado) collection. This lets you set and return parameter values without having to call the [Refresh](refresh-method-ado) method on the \*\*Parameters\*\* collection to retrieve the parameter information from the provider, a potentially resource-intensive operation.
The \*\*Parameter\*\* object is not safe for scripting.
This section contains the following topic.
- [Parameter Object Properties, Methods, and Events](parameter-object-properties-methods-and-events)
