<!-- fonte: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbenumerator?view=netframework-4.8.1 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Reference
monikers:
- netstandard-2.0-pp
- netframework-2.0
- netframework-3.0
- netframework-3.5
- netframework-4.0
- netframework-4.5
- netframework-4.5.1
- netframework-4.5.2
- netframework-4.6
- netframework-4.6.1
- netframework-4.6.2
- netframework-4.7
- netframework-4.7.1
- netframework-4.7.2
- netframework-4.8
- netframework-4.8.1
- net-10.0-pp
- net-11.0-pp
defaultMoniker: net-11.0-pp
versioningType: Ranged
title: OleDbEnumerator Class (System.Data.OleDb) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbenumerator?view=net-11.0-pp
uid: System.Data.OleDb.OleDbEnumerator
namespace: System.Data.OleDb
breadcrumb\_path: /dotnet/breadcrumb/toc.json
recommendations: true
author: dotnet-bot
ms.author: dotnetcontent
ms.date: 2025-07-01T00:00:00.0000000Z
show\_latex: true
uhfHeaderId: MSDocsHeader-DotNet
apiPlatform: dotnet
ms.topic: reference
ms.service: dotnet-api
products:
- https://authoring-docs-microsoft.poolparty.biz/devrel/7696cda6-0510-47f6-8302-71bb5d2e28cf
feedback\_system: OpenSource
feedback\_product\_url: https://aka.ms/feedback/report?space=61
feedback\_help\_link\_url: https://learn.microsoft.com/answers/tags/97/dotnet
feedback\_help\_link\_type: get-help-at-qna
ms.subservice: system.data
api\_name:
- System.Data.OleDb.OleDbEnumerator
api\_location:
- System.Data.dll
- System.Data.OleDb.dll
topic\_type:
- apiref
api\_type:
- Assembly
locale: en-us
document\_id: d3d7d2e7-e61f-8208-0c0d-73d567c950ed
document\_version\_independent\_id: d3d7d2e7-e61f-8208-0c0d-73d567c950ed
updated\_at: 2026-05-27T00:23:00.0000000Z
original\_content\_git\_url: https://github.com/dotnet/dotnet-api-docs-temp/blob/live/xml/System.Data.OleDb/OleDbEnumerator.xml
gitcommit: https://github.com/dotnet/dotnet-api-docs-temp/blob/e50d1863acf6da5affec9b315ba2ed88d12274e3/xml/System.Data.OleDb/OleDbEnumerator.xml
git\_commit\_id: e50d1863acf6da5affec9b315ba2ed88d12274e3
default\_moniker: net-11.0-pp
site\_name: Docs
depot\_name: Learn.dotnet-api-docs-temp
page\_type: dotnet
page\_kind: class
ms.assetid: System.Data.OleDb.OleDbEnumerator
description: 'Provides a mechanism for enumerating all available OLE DB providers within the local network. '
toc\_rel: dotnet-api-docs-temp/\_splitted/system.data.oledb/toc.json
search.mshattr.devlang: csharp vb fsharp cpp
asset\_id: api/system.data.oledb.oledbenumerator
moniker\_range\_name: 4e81f48678d316d605fd05a56d9e98af
monikers:
- netstandard-2.0-pp
- netframework-2.0
- netframework-3.0
- netframework-3.5
- netframework-4.0
- netframework-4.5
- netframework-4.5.1
- netframework-4.5.2
- netframework-4.6
- netframework-4.6.1
- netframework-4.6.2
- netframework-4.7
- netframework-4.7.1
- netframework-4.7.2
- netframework-4.8
- netframework-4.8.1
- net-10.0-pp
- net-11.0-pp
item\_type: Content
source\_path: xml/System.Data.OleDb/OleDbEnumerator.xml
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/7696cda6-0510-47f6-8302-71bb5d2e28cf
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/69c76c32-967e-4c65-b89a-74cc527db725
platformId: 2815eda9-fc27-c7a5-8bf6-2914c6a1bb16
---
# OleDbEnumerator Class
## Definition
- Namespace:
- [System.Data.OleDb](system.data.oledb)
- Assembly:
- System.Data.OleDb.dll
- Assembly:
- System.Data.dll
- Package:
- System.Data.OleDb v11.0.0-rc.1.26425.128
- Source:
- [OleDbEnumerator.cs](https://github.com/dotnet/dotnet/blob/3551975be08744f0418857c5bed8ab1545c5dd47/src/runtime/src/libraries/System.Data.OleDb/src/OleDbEnumerator.cs)
- Source:
- [OleDbEnumerator.cs](https://github.com/dotnet/runtime/blob/81cabf2857a01351e5ab578947c7403a5b128ad1/src/libraries/System.Data.OleDb/src/OleDbEnumerator.cs)
- Source:
- [OleDbEnumerator.cs](https://github.com/dotnet/runtime/blob/990ebf52fc408ca45929fd176d2740675a67fab8/src/libraries/System.Data.OleDb/src/OleDbEnumerator.cs)
- Source:
- [OleDbEnumerator.cs](https://github.com/dotnet/dotnet/blob/44525024595742ebe09023abe709df51de65009b/src/runtime/src/libraries/System.Data.OleDb/src/OleDbEnumerator.cs)
Provides a mechanism for enumerating all available OLE DB providers within the local network.
```cpp
public ref class OleDbEnumerator sealed
```
```csharp
public sealed class OleDbEnumerator
```
```fsharp
type OleDbEnumerator = class
```
```vb
Public NotInheritable Class OleDbEnumerator
```
- Inheritance
- [Object](system.object)
OleDbEnumerator
## Constructors
| Name | Description |
| --- | --- |
| [OleDbEnumerator()](system.data.oledb.oledbenumerator.-ctor#system-data-oledb-oledbenumerator-ctor) | Creates an instance of the [OleDbEnumerator](system.data.oledb.oledbenumerator) class. |
## Methods
| Name | Description |
| --- | --- |
| [Equals(Object)](system.object.equals#system-object-equals%28system-object%29) | Determines whether the specified object is equal to the current object.  
 (Inherited from [Object](system.object)) |
| [GetElements()](system.data.oledb.oledbenumerator.getelements#system-data-oledb-oledbenumerator-getelements) | Retrieves a [DataTable](system.data.datatable) that contains information about all visible OLE DB providers. |
| [GetEnumerator(Type)](system.data.oledb.oledbenumerator.getenumerator#system-data-oledb-oledbenumerator-getenumerator%28system-type%29) | Uses a specific OLE DB enumerator to return an [OleDbDataReader](system.data.oledb.oledbdatareader) that contains information about the currently installed OLE DB providers, without requiring an instance of the [OleDbEnumerator](system.data.oledb.oledbenumerator) class. |
| [GetHashCode()](system.object.gethashcode#system-object-gethashcode) | Serves as the default hash function.  
 (Inherited from [Object](system.object)) |
| [GetRootEnumerator()](system.data.oledb.oledbenumerator.getrootenumerator#system-data-oledb-oledbenumerator-getrootenumerator) | Returns an [OleDbDataReader](system.data.oledb.oledbdatareader) that contains information about the currently installed OLE DB providers, without requiring an instance of the [OleDbEnumerator](system.data.oledb.oledbenumerator) class. |
| [GetType()](system.object.gettype#system-object-gettype) | Gets the [Type](system.type) of the current instance.  
 (Inherited from [Object](system.object)) |
| [MemberwiseClone()](system.object.memberwiseclone#system-object-memberwiseclone) | Creates a shallow copy of the current [Object](system.object).  
 (Inherited from [Object](system.object)) |
| [ToString()](system.object.tostring#system-object-tostring) | Returns a string that represents the current object.  
 (Inherited from [Object](system.object)) |
## Applies to
