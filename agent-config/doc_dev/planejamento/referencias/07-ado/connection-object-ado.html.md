<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver16
versioningType: Ranged
title: Connection Object (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-ver16
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
description: Connection Object (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 49935499-f70c-eb0c-fb15-89d10d4f875f
document\_version\_independent\_id: 3caeabc8-feca-edc8-1ee1-4e95f56bc636
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/connection-object-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/connection-object-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver16
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 620
asset\_id: ado/reference/ado-api/connection-object-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/connection-object-ado.md
platformId: 3ec52c71-b495-9d64-afb0-af824f65aa24
---
# Connection Object (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Represents an open connection to a data source.
## Remarks
A \*\*Connection\*\* object represents a unique session with a data source. In a client/server database system, it may be equivalent to an actual network connection to the server. Depending on the functionality supported by the provider, some collections, methods, or properties of a \*\*Connection\*\* object may not be available.
With the collections, methods, and properties of a \*\*Connection\*\* object, you can do the following:
- Configure the connection before opening it with the [ConnectionString](connectionstring-property-ado), [ConnectionTimeout](connectiontimeout-property-ado), and [Mode](mode-property-ado) properties. \*\*ConnectionString\*\* is the default property of the \*\*Connection\*\* object.
- Set the [CursorLocation](cursorlocation-property-ado) property to client to invoke the [Microsoft Cursor Service for OLE DB](../../guide/appendixes/microsoft-cursor-service-for-ole-db-ado-service-component), which supports batch updates.
- Set the default database for the connection with the [DefaultDatabase](defaultdatabase-property) property.
- Set the level of isolation for the transactions opened on the connection with the [IsolationLevel](isolationlevel-property) property.
- Specify an OLE DB provider with the [Provider](provider-property-ado) property.
- Establish, and later break, the physical connection to the data source with the [Open](open-method-ado-connection) and [Close](close-method-ado) methods.
- Execute a command on the connection with the [Execute](execute-method-ado-connection) method and configure the execution with the [CommandTimeout](commandtimeout-property-ado) property.
Note
To execute a query without using a Command object, pass a query string to the \*\*Execute\*\* method of a \*\*Connection\*\* object. However, a [Command](command-object-ado) object is required when you want to persist the command text and re-execute it, or use query parameters.
- Manage transactions on the open connection, including nested transactions if the provider supports them, with the [BeginTrans](begintrans-committrans-and-rollbacktrans-methods-ado), [CommitTrans](begintrans-committrans-and-rollbacktrans-methods-ado), and [RollbackTrans](begintrans-committrans-and-rollbacktrans-methods-ado) methods and the [Attributes](attributes-property-ado) property.
- Examine errors returned from the data source with the [Errors](errors-collection-ado) collection.
- Read the version from the ADO implementation used with the [Version](version-property-ado) property.
- Obtain schema information about your database with the [OpenSchema](openschema-method) method.
You can create \*\*Connection\*\* objects independently of any other previously defined object.
You can execute named commands or stored procedures as if they were native methods on a \*\*Connection\*\* object, as shown in the next section. When a named command has the same name as that of a stored procedure, invoke the "native method call" on a \*\*Connection\*\* object always execute the named command instead of the store procedure.
Note
Do not use this feature (calling a named command or stored procedure as if it were a native method on the \*\*Connection\*\* object) in a Microsoft .NET Framework application, because the underlying implementation of the feature conflicts with the way the .NET Framework interoperates with COM.
## Execute a command as a native method of a Connection object
To execute a command, give the command a name using the \*\*Command\*\* object [Name](name-property-ado) property. Set the \*\*ActiveConnection\*\* property of the \*\*Command\*\* object to the connection. Then issue a statement where the command name is used as if it were a method on the \*\*Connection\*\* object, followed by any parameters, and a \*\*Recordset\*\* object if any rows are returned. Set the \*\*Recordset\*\* properties to customize the resulting \*\*Recordset\*\*. For example:
```
Dim cnn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
...
cnn.Open "..."
cmd.Name = "yourCommandName"
cmd.ActiveConnection = cnn
...
'Your command name, any parameters, and an optional Recordset.
cnn. "parameter", rst
```
## Execute a stored procedure as a native method of a Connection object
To execute a stored procedure, issue a statement where the stored procedure name is used as if it were a method on the \*\*Connection\*\* object, followed by any parameters. ADO will make a "best guess" of parameter types. For example:
```
Dim cnn As New ADODB.Connection
...
'Your stored procedure name and any parameters.
cnn. "parameter"
```
The \*\*Connection\*\* object is safe for scripting.
This section contains the following topic.
- [Connection Object Properties, Methods, and Events](connection-object-properties-methods-and-events)
---
## Other Supported Versions
- [sql-server-ver16](https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-ver16&accept=text/markdown)
