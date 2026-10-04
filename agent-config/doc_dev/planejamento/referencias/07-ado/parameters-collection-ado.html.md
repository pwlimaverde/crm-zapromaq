<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameters-collection-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver16
versioningType: Ranged
title: Parameters Collection (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameters-collection-ado?view=sql-server-ver16
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
description: Parameters Collection (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 457dee94-0548-c73f-8930-504f782637e5
document\_version\_independent\_id: 93ccdb12-8786-a78e-2606-4d90a3c44e44
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameters-collection-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameters-collection-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver16
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 602
asset\_id: ado/reference/ado-api/parameters-collection-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/parameters-collection-ado.md
platformId: 8f462278-42d7-a141-082d-f55ef37242cf
---
# Parameters Collection (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Contains all the [Parameter](parameter-object) objects of a [Command](command-object-ado) object.
## Remarks
A \*\*Command\*\* object has a \*\*Parameters\*\* collection made up of \*\*Parameter\*\* objects.
Using the [Refresh](refresh-method-ado) method on a \*\*Command\*\* object's \*\*Parameters\*\* collection retrieves provider parameter information for the stored procedure or parameterized query specified in the \*\*Command\*\* object. Some providers do not support stored procedure calls or parameterized queries; calling the \*\*Refresh\*\* method on the \*\*Parameters\*\* collection when using such a provider will return an error.
If you have not defined your own \*\*Parameter\*\* objects and you access the \*\*Parameters\*\* collection before calling the \*\*Refresh\*\* method, ADO will automatically call the method and populate the collection for you.
You can minimize calls to the provider to improve performance if you know the properties of the parameters associated with the stored procedure or parameterized query you wish to call. Use the [CreateParameter](createparameter-method-ado) method to create \*\*Parameter\*\* objects with the appropriate property settings and use the [Append](append-method-ado) method to add them to the \*\*Parameters\*\* collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must manually populate the \*\*Parameters\*\* collection using this method to be able to use parameters at all. Use the [Delete](delete-method-ado-parameters-collection) method to remove \*\*Parameter\*\* objects from the \*\*Parameters\*\* collection if necessary.
The objects in the \*\*Parameters\*\* collection of a \*\*Recordset\*\* go out of scope (therefore becoming unavailable) when the \*\*Recordset\*\* is closed.
When calling a stored procedure with \*\*Command\*\*, the return value/output parameter of a stored procedure is retrieved as follows:
1. When calling a stored procedure that has no parameters, the \*\*Refresh\*\* method on the \*\*Parameters\*\* collection should be called before calling the \*\*Execute\*\* method on the \*\*Command\*\* object.
2. When calling a stored procedure with parameters and explicitly appending a parameter to the \*\*Parameters\*\* collection with \*\*Append\*\*, the return value/output parameter should be appended to the \*\*Parameters\*\* collection. The return value must first be appended to the \*\*Parameters\*\* collection. Use \*\*Append\*\* to add the other parameters into the \*\*Parameters\*\* collection in the order of definition. For example, the stored procedure SPWithParam has two parameters. The first parameter, \*InParam\*, is an input parameter defined as adVarChar (20), and the second parameter, \*OutParam\*, is an output parameter defined as adVarChar (20). You can retrieve the return value/output parameter with the following code.
```vb
' Open Connection Conn
set ccmd = CreateObject("ADODB.Command")
ccmd.Activeconnection= Conn
ccmd.CommandText="SPWithParam"
ccmd.commandType = 4 'adCmdStoredProc
ccmd.parameters.Append ccmd.CreateParameter(, adInteger, adParamReturnValue, , NULL) ' return value
ccmd.parameters.Append ccmd.CreateParameter("InParam", adVarChar, adParamInput, 20, "hello world") ' input parameter
ccmd.parameters.Append ccmd.CreateParameter("OutParam", adVarChar, adParamOutput, 20, NULL) ' output parameter
ccmd.execute()
' Access ccmd.parameters(0) as return value of this stored procedure
' Access ccmd.parameters("OutParam") as the output parameter of this stored procedure.
```
3. When calling a stored procedure with parameters and configuring the parameters by calling the \*\*Item\*\* method on the \*\*Parameters\*\* collection, the return value/output parameter of the stored procedure can be retrieved from the \*\*Parameters\*\* collection. For example, the stored procedure SPWithParam has two parameters. The first parameter, \*InParam\*, is an input parameter defined as adVarChar (20), and the second parameter, \*OutParam\*, is an output parameter defined as adVarChar (20). You can retrieve the return value/output parameter with the following code.
```vb
' Open Connection Conn
set ccmd = CreateObject("ADODB.Command")
ccmd.Activeconnection= Conn
ccmd.CommandText="SPWithParam"
ccmd.commandType = 4 'adCmdStoredProc
ccmd.parameters.Item("InParam").value = "hello world" ' input parameter
ccmd.execute()
' Access ccmd.parameters(0) as return value of stored procedure
' Access ccmd.parameters(2) or ccmd.parameters("OutParam") as the output parameter.
```
This section contains the following topic.
- [Parameters Collection Properties, Methods, and Events](parameters-collection-properties-methods-and-events)
---
## Other Supported Versions
- [sql-server-ver16](https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/parameters-collection-ado?view=sql-server-ver16&accept=text/markdown)
