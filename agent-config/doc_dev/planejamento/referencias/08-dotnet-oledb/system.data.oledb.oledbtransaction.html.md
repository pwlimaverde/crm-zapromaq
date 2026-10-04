<!-- fonte: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbtransaction?view=netframework-4.8.1 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Reference
monikers:
- netframework-1.1
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
title: OleDbTransaction Class (System.Data.OleDb) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbtransaction?view=net-11.0-pp
uid: System.Data.OleDb.OleDbTransaction
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
- System.Data.OleDb.OleDbTransaction
api\_location:
- System.Data.dll
- System.Data.OleDb.dll
topic\_type:
- apiref
api\_type:
- Assembly
locale: en-us
document\_id: f4b6fb16-aaf9-d473-4d62-c0ef7596ebfa
document\_version\_independent\_id: f4b6fb16-aaf9-d473-4d62-c0ef7596ebfa
updated\_at: 2026-07-01T18:02:00.0000000Z
original\_content\_git\_url: https://github.com/dotnet/dotnet-api-docs-temp/blob/live/xml/System.Data.OleDb/OleDbTransaction.xml
gitcommit: https://github.com/dotnet/dotnet-api-docs-temp/blob/5370c326c6650b054fdaefc78fa9d4d44a5bc53e/xml/System.Data.OleDb/OleDbTransaction.xml
git\_commit\_id: 5370c326c6650b054fdaefc78fa9d4d44a5bc53e
default\_moniker: net-11.0-pp
site\_name: Docs
depot\_name: Learn.dotnet-api-docs-temp
page\_type: dotnet
page\_kind: class
ms.assetid: System.Data.OleDb.OleDbTransaction
description: 'Represents an SQL transaction to be made at a data source. This class cannot be inherited. '
toc\_rel: dotnet-api-docs-temp/\_splitted/system.data.oledb/toc.json
search.mshattr.devlang: csharp vb fsharp cpp
asset\_id: api/system.data.oledb.oledbtransaction
moniker\_range\_name: b59157a2302b2e4e771023712a03c52e
monikers:
- netframework-1.1
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
source\_path: xml/System.Data.OleDb/OleDbTransaction.xml
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/7696cda6-0510-47f6-8302-71bb5d2e28cf
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/69c76c32-967e-4c65-b89a-74cc527db725
platformId: fc155ad5-56ed-9eca-d159-e9a5fce34f4e
---
# OleDbTransaction Class
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
- [OleDbTransaction.cs](https://github.com/dotnet/dotnet/blob/3551975be08744f0418857c5bed8ab1545c5dd47/src/runtime/src/libraries/System.Data.OleDb/src/OleDbTransaction.cs)
- Source:
- [OleDbTransaction.cs](https://github.com/dotnet/runtime/blob/81cabf2857a01351e5ab578947c7403a5b128ad1/src/libraries/System.Data.OleDb/src/OleDbTransaction.cs)
- Source:
- [OleDbTransaction.cs](https://github.com/dotnet/runtime/blob/990ebf52fc408ca45929fd176d2740675a67fab8/src/libraries/System.Data.OleDb/src/OleDbTransaction.cs)
- Source:
- [OleDbTransaction.cs](https://github.com/dotnet/dotnet/blob/44525024595742ebe09023abe709df51de65009b/src/runtime/src/libraries/System.Data.OleDb/src/OleDbTransaction.cs)
Represents an SQL transaction to be made at a data source. This class cannot be inherited.
```cpp
public ref class OleDbTransaction sealed : System::Data::Common::DbTransaction
```
```cpp
public ref class OleDbTransaction sealed : MarshalByRefObject, IDisposable, System::Data::IDbTransaction
```
```csharp
[System.Diagnostics.CodeAnalysis.RequiresDynamicCode("OleDbConnection is not AOT-compatible.")]
public sealed class OleDbTransaction : System.Data.Common.DbTransaction
```
```csharp
public sealed class OleDbTransaction : MarshalByRefObject, IDisposable, System.Data.IDbTransaction
```
```csharp
public sealed class OleDbTransaction : System.Data.Common.DbTransaction
```
```fsharp
[]
type OleDbTransaction = class
inherit DbTransaction
```
```fsharp
type OleDbTransaction = class
inherit MarshalByRefObject
interface IDbTransaction
interface IDisposable
```
```fsharp
type OleDbTransaction = class
inherit DbTransaction
```
```vb
Public NotInheritable Class OleDbTransaction
Inherits DbTransaction
```
```vb
Public NotInheritable Class OleDbTransaction
Inherits MarshalByRefObject
Implements IDbTransaction, IDisposable
```
- Inheritance
- [Object](system.object)
[DbTransaction](system.data.common.dbtransaction)
OleDbTransaction
- Inheritance
- [Object](system.object)
[MarshalByRefObject](system.marshalbyrefobject)
[DbTransaction](system.data.common.dbtransaction)
OleDbTransaction
- Inheritance
- [Object](system.object)
[MarshalByRefObject](system.marshalbyrefobject)
OleDbTransaction
- Attributes
- [RequiresDynamicCodeAttribute](system.diagnostics.codeanalysis.requiresdynamiccodeattribute)
- Implements
- [IDbTransaction](system.data.idbtransaction)[IDisposable](system.idisposable)
## Remarks
The application creates an [OleDbTransaction](system.data.oledb.oledbtransaction) object by calling [BeginTransaction](system.data.oledb.oledbconnection.begintransaction) on the [OleDbConnection](system.data.oledb.oledbconnection) object. All subsequent operations associated with the transaction (for example, committing or aborting the transaction), are performed on the [OleDbTransaction](system.data.oledb.oledbtransaction) object.
## Properties
| Name | Description |
| --- | --- |
| [Connection](system.data.oledb.oledbtransaction.connection#system-data-oledb-oledbtransaction-connection) | Gets the [OleDbConnection](system.data.oledb.oledbconnection) object associated with the transaction, or `null` if the transaction is no longer valid. |
| [DbConnection](system.data.common.dbtransaction.dbconnection#system-data-common-dbtransaction-dbconnection) | When overridden in a derived class, gets the [DbConnection](system.data.common.dbconnection) object associated with the transaction.  
 (Inherited from [DbTransaction](system.data.common.dbtransaction)) |
| [IsolationLevel](system.data.oledb.oledbtransaction.isolationlevel#system-data-oledb-oledbtransaction-isolationlevel) | Specifies the [IsolationLevel](system.data.isolationlevel) for this transaction. |
## Methods
| Name | Description |
| --- | --- |
| [Begin()](system.data.oledb.oledbtransaction.begin#system-data-oledb-oledbtransaction-begin) | Initiates a nested database transaction. |
| [Begin(IsolationLevel)](system.data.oledb.oledbtransaction.begin#system-data-oledb-oledbtransaction-begin%28system-data-isolationlevel%29) | Initiates a nested database transaction and specifies the isolation level to use for the new transaction. |
| [Commit()](system.data.oledb.oledbtransaction.commit#system-data-oledb-oledbtransaction-commit) | Commits the database transaction. |
| [CreateObjRef(Type)](system.marshalbyrefobject.createobjref#system-marshalbyrefobject-createobjref%28system-type%29) | Creates an object that contains all the relevant information required to generate a proxy used to communicate with a remote object.  
 (Inherited from [MarshalByRefObject](system.marshalbyrefobject)) |
| [Dispose()](system.data.common.dbtransaction.dispose#system-data-common-dbtransaction-dispose) | Releases the unmanaged resources used by the [DbTransaction](system.data.common.dbtransaction).  
 (Inherited from [DbTransaction](system.data.common.dbtransaction)) |
| [Dispose(Boolean)](system.data.common.dbtransaction.dispose#system-data-common-dbtransaction-dispose%28system-boolean%29) | Releases the unmanaged resources used by the [DbTransaction](system.data.common.dbtransaction) and optionally releases the managed resources.  
 (Inherited from [DbTransaction](system.data.common.dbtransaction)) |
| [Equals(Object)](system.object.equals#system-object-equals%28system-object%29) | Determines whether the specified object is equal to the current object.  
 (Inherited from [Object](system.object)) |
| [Finalize()](system.data.oledb.oledbtransaction.finalize#system-data-oledb-oledbtransaction-finalize) | Allows an object to try to free resources and perform other cleanup operations before it is reclaimed by garbage collection. |
| [GetHashCode()](system.object.gethashcode#system-object-gethashcode) | Serves as the default hash function.  
 (Inherited from [Object](system.object)) |
| [GetLifetimeService()](system.marshalbyrefobject.getlifetimeservice#system-marshalbyrefobject-getlifetimeservice) | ::: moniker range=" net-10.0 net-11.0 net-5.0 net-6.0 net-7.0 net-8.0 net-9.0 "  
\*\*Obsolete.\*\*  
  
::: moniker-end  
  
Retrieves the current lifetime service object that controls the lifetime policy for this instance.  
 (Inherited from [MarshalByRefObject](system.marshalbyrefobject)) |
| [GetType()](system.object.gettype#system-object-gettype) | Gets the [Type](system.type) of the current instance.  
 (Inherited from [Object](system.object)) |
| [InitializeLifetimeService()](system.marshalbyrefobject.initializelifetimeservice#system-marshalbyrefobject-initializelifetimeservice) | ::: moniker range=" net-10.0 net-11.0 net-5.0 net-6.0 net-7.0 net-8.0 net-9.0 "  
\*\*Obsolete.\*\*  
  
::: moniker-end  
  
Obtains a lifetime service object to control the lifetime policy for this instance.  
 (Inherited from [MarshalByRefObject](system.marshalbyrefobject)) |
| [MemberwiseClone()](system.object.memberwiseclone#system-object-memberwiseclone) | Creates a shallow copy of the current [Object](system.object).  
 (Inherited from [Object](system.object)) |
| [MemberwiseClone(Boolean)](system.marshalbyrefobject.memberwiseclone#system-marshalbyrefobject-memberwiseclone%28system-boolean%29) | Creates a shallow copy of the current [MarshalByRefObject](system.marshalbyrefobject) object.  
 (Inherited from [MarshalByRefObject](system.marshalbyrefobject)) |
| [Rollback()](system.data.oledb.oledbtransaction.rollback#system-data-oledb-oledbtransaction-rollback) | Rolls back a transaction from a pending state. |
| [ToString()](system.object.tostring#system-object-tostring) | Returns a string that represents the current object.  
 (Inherited from [Object](system.object)) |
## Explicit Interface Implementations
| Name | Description |
| --- | --- |
| [IDbTransaction.Connection](system.data.common.dbtransaction.system-data-idbtransaction-connection#system-data-common-dbtransaction-system-data-idbtransaction-connection) | Gets the [DbConnection](system.data.common.dbconnection) object associated with the transaction, or a null reference if the transaction is no longer valid.  
 (Inherited from [DbTransaction](system.data.common.dbtransaction)) |
| [IDisposable.Dispose()](system.data.oledb.oledbtransaction.system-idisposable-dispose#system-data-oledb-oledbtransaction-system-idisposable-dispose) | Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources. |
## Applies to
