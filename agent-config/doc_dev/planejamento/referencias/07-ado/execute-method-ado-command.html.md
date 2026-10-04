<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/execute-method-ado-command?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver16
versioningType: Ranged
title: Execute Method (ADO Command) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/execute-method-ado-command?view=sql-server-ver16
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
description: Execute Method (ADO Command)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: f84675bd-303e-fd57-017b-8d50b214381c
document\_version\_independent\_id: ba66b7ae-58ee-f723-f39d-c0395193a891
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/execute-method-ado-command.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/execute-method-ado-command.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver16
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 661
asset\_id: ado/reference/ado-api/execute-method-ado-command
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/execute-method-ado-command.md
platformId: 131c7591-d02d-fc9e-ba3e-c42eab820b93
---
# Execute Method (ADO Command) - ActiveX Data Objects (ADO) | Microsoft Learn
Executes the query, SQL statement, or stored procedure specified in the [CommandText](commandtext-property-ado) or [CommandStream](commandstream-property-ado) property of the [Command object](command-object-ado).
## Syntax
```
Set recordset = command.Execute( RecordsAffected, Parameters, Options )
```
## Return Value
Returns a [Recordset](recordset-object-ado) object reference, a stream, or \*\*Nothing\*\*.
#### Parameters
\*RecordsAffected\* Optional. A \*\*Long\*\* variable to which the provider returns the number of records that the operation affected. The \*RecordsAffected\* parameter applies only for action queries or stored procedures. \*RecordsAffected\* does not return the number of records returned by a result-returning query or stored procedure. To obtain this information, use the [RecordCount](recordcount-property-ado) property. The \*\*Execute\*\* method will not return the correct information when used with \*\*adAsyncExecute\*\*, simply because when a command is executed asynchronously, the number of records affected may not yet be known at the time the method returns.
\*Parameters\* Optional. A \*\*Variant\*\* array of parameter values used in conjunction with the input string or stream specified in \*\*CommandText\*\* or \*\*CommandStream\*\*. (Output parameters will not return correct values when passed in this argument.)
\*Options\* Optional. A \*\*Long\*\* value that indicates how the provider should evaluate the [CommandText](commandtext-property-ado) or the [CommandStream](commandstream-property-ado) property of the [Command](command-object-ado) object. Can be a bitmask value made using [CommandTypeEnum](commandtypeenum) and/or [ExecuteOptionEnum](executeoptionenum) values. For example, you could use \*\*adCmdText\*\* and \*\*adExecuteNoRecords\*\* in combination if you want to have ADO evaluate the value of the \*\*CommandText\*\* property as text, and indicate that the command should discard and not return any records that might be generated when the command text executes.
Note
Use the \*\*ExecuteOptionEnum\*\* value \*\*adExecuteNoRecords\*\* to improve performance by minimizing internal processing. If \*\*adExecuteStream\*\* was specified, the options \*\*adAsyncFetch\*\* and \*\*adAsynchFetchNonBlocking\*\* are ignored. Do not use the \*\*CommandTypeEnum\*\* values of \*\*adCmdFile\*\* or \*\*adCmdTableDirect\*\* with \*\*Execute\*\*. These values can only be used as options with the [Open](open-method-ado-recordset) and [Requery](requery-method) methods of a \*\*Recordset\*\*.
## Remarks
Using the \*\*Execute\*\* method on a \*\*Command\*\* object executes the query specified in the \*\*CommandText\*\* property or \*\*CommandStream\*\* property of the object.
Results are returned in a \*\*Recordset\*\* (by default) or as a stream of binary information. To obtain a binary stream, specify \*\*adExecuteStream\*\* in \*Options\*, then supply a stream by setting \*\*Command.Properties("Output Stream")\*\*. An ADO \*\*Stream\*\* object can be specified to receive the results, or another stream object such as the IIS Response object can be specified. If no stream was specified before calling \*\*Execute\*\* with \*\*adExecuteStream\*\*, an error occurs. The position of the stream on return from \*\*Execute\*\* is provider specific.
If the command is not intended to return results (for example, a SQL UPDATE query) the provider returns \*\*Nothing\*\* as long as the option \*\*adExecuteNoRecords\*\* is specified; otherwise Execute returns a closed \*\*Recordset\*\*. Some application languages allow you to ignore this return value if no \*\*Recordset\*\* is desired.
\*\*Execute\*\* raises an error if the user specifies a value for \*\*CommandStream\*\* when the \*\*CommandType\*\* is \*\*adCmdStoredProc\*\*, \*\*adCmdTable\*\*, or \*\*adCmdTableDirect\*\*.
If the query has parameters, the current values for the \*\*Command\*\* object's parameters are used unless you override these with parameter values passed with the \*\*Execute\*\* call. You can override a subset of the parameters by omitting new values for some of the parameters when calling the \*\*Execute\*\* method. The order in which you specify the parameters is the same order in which the method passes them. For example, if there were four (or more) parameters and you wanted to pass new values for only the first and fourth parameters, you would pass `Array(var1,,,var4)` as the \*Parameters\* argument.
Note
Output parameters will not return correct values when passed in the \*Parameters\* argument.
An [ExecuteComplete](executecomplete-event-ado) event will be issued when this operation concludes.
Note
When issuing commands containing URLs, those using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](../../guide/appendixes/microsoft-ole-db-provider-for-internet-publishing). For more information, see [Absolute and Relative URLs](../../guide/data/absolute-and-relative-urls).
## Applies To
[Command Object (ADO)](command-object-ado)
---
## Other Supported Versions
- [sql-server-ver16](https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/execute-method-ado-command?view=sql-server-ver16&accept=text/markdown)
