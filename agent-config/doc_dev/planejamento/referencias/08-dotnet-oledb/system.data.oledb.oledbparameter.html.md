<!-- fonte: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbparameter?view=netframework-4.8.1 | obtido em: 2026-09-19 | convertido de HTML -->

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
title: OleDbParameter Class (System.Data.OleDb) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbparameter?view=net-11.0-pp
uid: System.Data.OleDb.OleDbParameter
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
- System.Data.OleDb.OleDbParameter
api\_location:
- System.Data.dll
- System.Data.OleDb.dll
topic\_type:
- apiref
api\_type:
- Assembly
locale: en-us
document\_id: b9cc319f-2005-2491-8ea4-3272e35f1b4e
document\_version\_independent\_id: b9cc319f-2005-2491-8ea4-3272e35f1b4e
updated\_at: 2026-07-01T18:02:00.0000000Z
original\_content\_git\_url: https://github.com/dotnet/dotnet-api-docs-temp/blob/live/xml/System.Data.OleDb/OleDbParameter.xml
gitcommit: https://github.com/dotnet/dotnet-api-docs-temp/blob/5370c326c6650b054fdaefc78fa9d4d44a5bc53e/xml/System.Data.OleDb/OleDbParameter.xml
git\_commit\_id: 5370c326c6650b054fdaefc78fa9d4d44a5bc53e
default\_moniker: net-11.0-pp
site\_name: Docs
depot\_name: Learn.dotnet-api-docs-temp
page\_type: dotnet
page\_kind: class
ms.assetid: System.Data.OleDb.OleDbParameter
description: 'Represents a parameter to an OleDbCommand and optionally its mapping to a DataSet column. This class cannot be inherited. '
toc\_rel: dotnet-api-docs-temp/\_splitted/system.data.oledb/toc.json
search.mshattr.devlang: csharp vb fsharp cpp
asset\_id: api/system.data.oledb.oledbparameter
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
source\_path: xml/System.Data.OleDb/OleDbParameter.xml
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/7696cda6-0510-47f6-8302-71bb5d2e28cf
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/69c76c32-967e-4c65-b89a-74cc527db725
platformId: 3c6f2a1a-b62d-ffe6-830d-4d0c7ab92114
---
# OleDbParameter Class
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
- [OleDbParameter.cs](https://github.com/dotnet/dotnet/blob/3551975be08744f0418857c5bed8ab1545c5dd47/src/runtime/src/libraries/System.Data.OleDb/src/OleDbParameter.cs)
- Source:
- [OleDbParameter.cs](https://github.com/dotnet/runtime/blob/81cabf2857a01351e5ab578947c7403a5b128ad1/src/libraries/System.Data.OleDb/src/OleDbParameter.cs)
- Source:
- [OleDbParameter.cs](https://github.com/dotnet/runtime/blob/990ebf52fc408ca45929fd176d2740675a67fab8/src/libraries/System.Data.OleDb/src/OleDbParameter.cs)
- Source:
- [OleDbParameter.cs](https://github.com/dotnet/dotnet/blob/44525024595742ebe09023abe709df51de65009b/src/runtime/src/libraries/System.Data.OleDb/src/OleDbParameter.cs)
Represents a parameter to an [OleDbCommand](system.data.oledb.oledbcommand) and optionally its mapping to a [DataSet](system.data.dataset) column. This class cannot be inherited.
```cpp
public ref class OleDbParameter sealed : System::Data::Common::DbParameter, ICloneable
```
```cpp
public ref class OleDbParameter sealed : MarshalByRefObject, ICloneable, System::Data::IDbDataParameter
```
```csharp
[System.ComponentModel.TypeConverter(typeof(System.Data.OleDb.OleDbParameter+OleDbParameterConverter))]
public sealed class OleDbParameter : System.Data.Common.DbParameter, ICloneable
```
```csharp
[System.ComponentModel.TypeConverter(typeof(System.Data.OleDb.OleDbParameterConverter))]
public sealed class OleDbParameter : MarshalByRefObject, ICloneable, System.Data.IDbDataParameter
```
```fsharp
[]
type OleDbParameter = class
inherit DbParameter
interface IDataParameter
interface IDbDataParameter
interface ICloneable
```
```fsharp
[]
type OleDbParameter = class
inherit MarshalByRefObject
interface IDbDataParameter
interface IDataParameter
interface ICloneable
```
```fsharp
[]
type OleDbParameter = class
inherit DbParameter
interface ICloneable
interface IDbDataParameter
interface IDataParameter
```
```vb
Public NotInheritable Class OleDbParameter
Inherits DbParameter
Implements ICloneable
```
```vb
Public NotInheritable Class OleDbParameter
Inherits MarshalByRefObject
Implements ICloneable, IDbDataParameter
```
- Inheritance
- [Object](system.object)
[DbParameter](system.data.common.dbparameter)
OleDbParameter
- Inheritance
- [Object](system.object)
[MarshalByRefObject](system.marshalbyrefobject)
[DbParameter](system.data.common.dbparameter)
OleDbParameter
- Inheritance
- [Object](system.object)
[MarshalByRefObject](system.marshalbyrefobject)
OleDbParameter
- Attributes
- [TypeConverterAttribute](system.componentmodel.typeconverterattribute)
- Implements
- [IDataParameter](system.data.idataparameter)[IDbDataParameter](system.data.idbdataparameter)[ICloneable](system.icloneable)
## Examples
The following example creates multiple instances of [OleDbParameter](system.data.oledb.oledbparameter) through the [OleDbParameterCollection](system.data.oledb.oledbparametercollection) collection within the [OleDbDataAdapter](system.data.oledb.oledbdataadapter). These parameters are used to select data from the data source and place the data in the [DataSet](system.data.dataset). This example assumes that a [DataSet](system.data.dataset) and an [OleDbDataAdapter](system.data.oledb.oledbdataadapter) have already been created by using the appropriate schema, commands, and connection.
```csharp
public DataSet GetDataSetFromAdapter(
DataSet dataSet, string connectionString, string queryString)
{
using (OleDbConnection connection =
new OleDbConnection(connectionString))
{
OleDbDataAdapter adapter =
new OleDbDataAdapter(queryString, connection);
// Set the parameters.
adapter.SelectCommand.Parameters.Add(
"@CategoryName", OleDbType.VarChar, 80).Value = "toasters";
adapter.SelectCommand.Parameters.Add(
"@SerialNum", OleDbType.Integer).Value = 239;
// Open the connection and fill the DataSet.
try
{
connection.Open();
adapter.Fill(dataSet);
}
catch (Exception ex)
{
Console.WriteLine(ex.Message);
}
// The connection is automatically closed when the
// code exits the using block.
}
return dataSet;
}
```
```vb
Public Function GetDataSetFromAdapter( \_
ByVal dataSet As DataSet, ByVal connectionString As String, \_
ByVal queryString As String) As DataSet
Using connection As New OleDbConnection(connectionString)
Dim adapter As New OleDbDataAdapter(queryString, connection)
' Set the parameters.
adapter.SelectCommand.Parameters.Add( \_
"@CategoryName", OleDbType.VarChar, 80).Value = "toasters"
adapter.SelectCommand.Parameters.Add( \_
"@SerialNum", OleDbType.Integer).Value = 239
' Open the connection and fill the DataSet.
Try
connection.Open()
adapter.Fill(dataSet)
Catch ex As Exception
Console.WriteLine(ex.Message)
End Try
' The connection is automatically closed when the
' code exits the Using block.
End Using
Return dataSet
End Function
```
## Remarks
The OLE DB.NET Framework Data Provider uses positional parameters that are marked with a question mark (?) instead of named parameters.
When querying an Oracle database using the Microsoft OLE DB Provider for Oracle (MSDAORA) and the OLE DB.NET Framework Data Provider, using the `LIKE` clause to query values in fixed-length fields may not return all expected matches. The reason is that when Oracle matches values for fixed-length fields in a `LIKE` clause, it matches the entire length of the string, including any padding trailing spaces. For example, if a table in an Oracle database contains a field named "Field1" that is defined as `char(3)`, and you enter the value "a" into a row of that table, the following code does not return the row.
```vb
Dim queryString As String = "SELECT \* FROM Table1 WHERE Field1 LIKE ?"
Dim command As OleDbCommand = New OleDbCommand(queryString, connection)
command.Parameters.Add("@p1", OleDbType.Char, 3).Value = "a"
Dim reader As OleDbDataReader = command.ExecuteReader()
```
```csharp
string queryString = "SELECT \* FROM Table1 WHERE Field1 LIKE ?";
OleDbCommand command = new OleDbCommand(queryString, connection);
command.Parameters.Add("@p1", OleDbType.Char, 3).Value = "a";
OleDbDataReader reader = command.ExecuteReader();
```
This is because Oracle stores the column value as "a " (padding "a", with trailing spaces, to the fixed field length of 3), which Oracle does not treat as a match for the parameter value of "a" in the case of a `LIKE` comparison of fixed-length fields.
To resolve this problem, append a percentage ("%") wildcard character to the parameter value (`"a%"`), or use an SQL `=` comparison instead.
## Constructors
| Name | Description |
| --- | --- |
| [OleDbParameter()](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class. |
| [OleDbParameter(String, Object)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-object%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name and the value of the new [OleDbParameter](system.data.oledb.oledbparameter). |
| [OleDbParameter(String, OleDbType, Int32, ParameterDirection, Boolean, Byte, Byte, String, DataRowVersion, Object)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-data-oledb-oledbtype-system-int32-system-data-parameterdirection-system-boolean-system-byte-system-byte-system-string-system-data-datarowversion-system-object%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name, data type, length, source column name, parameter direction, numeric precision, and other properties. |
| [OleDbParameter(String, OleDbType, Int32, ParameterDirection, Byte, Byte, String, DataRowVersion, Boolean, Object)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-data-oledb-oledbtype-system-int32-system-data-parameterdirection-system-byte-system-byte-system-string-system-data-datarowversion-system-boolean-system-object%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name, data type, length, source column name, parameter direction, numeric precision, and other properties. |
| [OleDbParameter(String, OleDbType, Int32, String)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-data-oledb-oledbtype-system-int32-system-string%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name, data type, length, and source column name. |
| [OleDbParameter(String, OleDbType, Int32)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-data-oledb-oledbtype-system-int32%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name, data type, and length. |
| [OleDbParameter(String, OleDbType)](system.data.oledb.oledbparameter.-ctor#system-data-oledb-oledbparameter-ctor%28system-string-system-data-oledb-oledbtype%29) | Initializes a new instance of the [OleDbParameter](system.data.oledb.oledbparameter) class that uses the parameter name and data type. |
## Properties
| Name | Description |
| --- | --- |
| [DbType](system.data.oledb.oledbparameter.dbtype#system-data-oledb-oledbparameter-dbtype) | Gets or sets the [DbType](system.data.dbtype) of the parameter. |
| [Direction](system.data.oledb.oledbparameter.direction#system-data-oledb-oledbparameter-direction) | Gets or sets a value that indicates whether the parameter is input-only, output-only, bidirectional, or a stored procedure return-value parameter. |
| [IsNullable](system.data.oledb.oledbparameter.isnullable#system-data-oledb-oledbparameter-isnullable) | Gets or sets a value that indicates whether the parameter accepts null values. |
| [OleDbType](system.data.oledb.oledbparameter.oledbtype#system-data-oledb-oledbparameter-oledbtype) | Gets or sets the [OleDbType](system.data.oledb.oledbtype) of the parameter. |
| [ParameterName](system.data.oledb.oledbparameter.parametername#system-data-oledb-oledbparameter-parametername) | Gets or sets the name of the [OleDbParameter](system.data.oledb.oledbparameter). |
| [Precision](system.data.oledb.oledbparameter.precision#system-data-oledb-oledbparameter-precision) | Gets or sets the maximum number of digits used to represent the [Value](system.data.oledb.oledbparameter.value#system-data-oledb-oledbparameter-value) property. |
| [Scale](system.data.oledb.oledbparameter.scale#system-data-oledb-oledbparameter-scale) | Gets or sets the number of decimal places to which [Value](system.data.oledb.oledbparameter.value#system-data-oledb-oledbparameter-value) is resolved. |
| [Size](system.data.oledb.oledbparameter.size#system-data-oledb-oledbparameter-size) | Gets or sets the maximum size, in bytes, of the data within the column. |
| [SourceColumn](system.data.oledb.oledbparameter.sourcecolumn#system-data-oledb-oledbparameter-sourcecolumn) | Gets or sets the name of the source column mapped to the [DataSet](system.data.dataset) and used for loading or returning the [Value](system.data.oledb.oledbparameter.value#system-data-oledb-oledbparameter-value). |
| [SourceColumnNullMapping](system.data.oledb.oledbparameter.sourcecolumnnullmapping#system-data-oledb-oledbparameter-sourcecolumnnullmapping) | Gets or sets a value that indicates whether the source column is nullable. This allows [DbCommandBuilder](system.data.common.dbcommandbuilder) to correctly generate Update statements for nullable columns. |
| [SourceVersion](system.data.oledb.oledbparameter.sourceversion#system-data-oledb-oledbparameter-sourceversion) | Gets or sets the [DataRowVersion](system.data.datarowversion) to use when you load [Value](system.data.oledb.oledbparameter.value#system-data-oledb-oledbparameter-value). |
| [Value](system.data.oledb.oledbparameter.value#system-data-oledb-oledbparameter-value) | Gets or sets the value of the parameter. |
## Methods
| Name | Description |
| --- | --- |
| [CreateObjRef(Type)](system.marshalbyrefobject.createobjref#system-marshalbyrefobject-createobjref%28system-type%29) | Creates an object that contains all the relevant information required to generate a proxy used to communicate with a remote object.  
 (Inherited from [MarshalByRefObject](system.marshalbyrefobject)) |
| [Equals(Object)](system.object.equals#system-object-equals%28system-object%29) | Determines whether the specified object is equal to the current object.  
 (Inherited from [Object](system.object)) |
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
| [ResetDbType()](system.data.oledb.oledbparameter.resetdbtype#system-data-oledb-oledbparameter-resetdbtype) | Resets the type associated with this [OleDbParameter](system.data.oledb.oledbparameter). |
| [ResetOleDbType()](system.data.oledb.oledbparameter.resetoledbtype#system-data-oledb-oledbparameter-resetoledbtype) | Resets the type associated with this [OleDbParameter](system.data.oledb.oledbparameter). |
| [ToString()](system.data.oledb.oledbparameter.tostring#system-data-oledb-oledbparameter-tostring) | Gets a string that contains the [ParameterName](system.data.oledb.oledbparameter.parametername#system-data-oledb-oledbparameter-parametername). |
## Explicit Interface Implementations
| Name | Description |
| --- | --- |
| [ICloneable.Clone()](system.data.oledb.oledbparameter.system-icloneable-clone#system-data-oledb-oledbparameter-system-icloneable-clone) | For a description of this member, see [Clone()](system.icloneable.clone#system-icloneable-clone). |
| [IDbDataParameter.Precision](system.data.common.dbparameter.system-data-idbdataparameter-precision#system-data-common-dbparameter-system-data-idbdataparameter-precision) | Indicates the precision of numeric parameters.  
 (Inherited from [DbParameter](system.data.common.dbparameter)) |
| [IDbDataParameter.Scale](system.data.common.dbparameter.system-data-idbdataparameter-scale#system-data-common-dbparameter-system-data-idbdataparameter-scale) | For a description of this member, see [Scale](system.data.idbdataparameter.scale#system-data-idbdataparameter-scale).  
 (Inherited from [DbParameter](system.data.common.dbparameter)) |
## Applies to
