<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/command-object-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver16
versioningType: Ranged
title: Command Object (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/command-object-ado?view=sql-server-ver16
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
description: Command Object (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: c6f95b71-b692-2580-af7c-6a9ac1ae3417
document\_version\_independent\_id: dce0c961-89c9-d0ca-0d00-547110a6463a
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/command-object-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/command-object-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver16
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 601
asset\_id: ado/reference/ado-api/command-object-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/command-object-ado.md
platformId: 63e63085-b665-9906-460b-0ebcea5eeb74
---
# Command Object (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Defines a specific command that you intend to execute against a data source.
## Remarks
Use a \*\*Command\*\* object to query a database and return records in a [Recordset](recordset-object-ado) object, to execute a bulk operation, or to manipulate the structure of a database. Depending on the functionality of the provider, some \*\*Command\*\* collections, methods, or properties may generate an error when they are referenced.
With the collections, methods, and properties of a \*\*Command\*\* object, you can do the following:
- Define the executable text of the command (for example, a SQL statement) with the [CommandText](commandtext-property-ado) property. Alternatively, for command or query structures other than simple strings (for example, XML template queries) define the command with the [CommandStream](commandstream-property-ado) property.
- Optionally, indicate the command dialect used in the \*\*CommandText\*\* or \*\*CommandStream\*\* with the [Dialect](dialect-property) property.
- Define parameterized queries or stored-procedure arguments with [Parameter](parameter-object) objects and the [Parameters](parameters-collection-ado) collection.
- Indicate whether parameter names should be passed to the provider with the [NamedParameters](namedparameters-property-ado) property.
- Execute a command and return a \*\*Recordset\*\* object if appropriate with the [Execute](execute-method-ado-command) method.
- Specify the type of command with the [CommandType](commandtype-property-ado) property prior to execution to optimize performance.
- Control whether the provider saves a prepared (or compiled) version of the command prior to execution with the [Prepared](prepared-property-ado) property.
- Set the number of seconds that a provider will wait for a command to execute with the [CommandTimeout](commandtimeout-property-ado) property.
- Associate an open connection with a \*\*Command\*\* object by setting its [ActiveConnection](activeconnection-property-ado) property.
- Set the [Name](name-property-ado) property to identify the \*\*Command\*\* object as a method on the associated [Connection](connection-object-ado) object.
- Pass a \*\*Command\*\* object to the [Source](source-property-ado-recordset) property of a \*\*Recordset\*\* to obtain data.
- Access provider-specific attributes with the [Properties](properties-collection-ado) collection.
Note
To execute a query without using a \*\*Command\*\* object, pass a query string to the [Execute](execute-method-ado-connection) method of a \*\*Connection\*\* object or to the [Open](open-method-ado-recordset) method of a \*\*Recordset\*\* object. However, a \*\*Command\*\* object is required when you want to persist the command text and re-execute it, or use query parameters.
To create a \*\*Command\*\* object independently of a previously defined \*\*Connection\*\* object, set its \*\*ActiveConnection\*\* property to a valid connection string. ADO still creates a \*\*Connection\*\* object, but it does not assign that object to an object variable. However, if you are associating multiple \*\*Command\*\* objects with the same connection, you should explicitly create and open a \*\*Connection\*\* object; this assigns the \*\*Connection\*\* object to an object variable. Make sure that the \*\*Connection\*\* object was opened successfully before you assign it to the \*\*ActiveConnection\*\* property of the \*\*Command\*\* object, because assigning a closed \*\*Connection\*\* object causes an error. If you do not set the \*\*ActiveConnection\*\* property of the \*\*Command\*\* object to this object variable, ADO creates a new \*\*Connection\*\* object for each \*\*Command\*\* object, even if you use the same connection string.
To execute a \*\*Command\*\*, call it by its [Name](name-property-ado) property on the associated \*\*Connection\*\* object. The \*\*Command\*\* must have its \*\*ActiveConnection\*\* property set to the \*\*Connection\*\* object. If the \*\*Command\*\* has parameters, pass their values as arguments to the method.
If two or more \*\*Command\*\* objects are executed on the same connection and either \*\*Command\*\* object is a stored procedure with output parameters, an error occurs. To execute each \*\*Command\*\* object, use separate connections or disconnect all other \*\*Command\*\* objects from the connection.
The \*\*Parameters\*\* collection is the default member of the \*\*Command\*\* object. As a result, the following two code statements are equivalent.
```
objCmd.Parameters.Item(0)
objCmd(0)
```
- The \*\*Command\*\* object is not safe for scripting.
This section contains the following topic.
- [Command Object Properties, Methods, and Events](command-object-properties-methods-and-events)
---
## Other Supported Versions
- [sql-server-ver16](https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/command-object-ado?view=sql-server-ver16&accept=text/markdown)
