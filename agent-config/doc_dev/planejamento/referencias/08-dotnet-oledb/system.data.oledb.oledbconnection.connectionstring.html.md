<!-- fonte: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbconnection.connectionstring?view=netframework-4.8.1 | obtido em: 2026-09-19 | convertido de HTML -->

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
title: OleDbConnection.ConnectionString Property (System.Data.OleDb) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbconnection.connectionstring?view=net-11.0-pp
uid: System.Data.OleDb.OleDbConnection.ConnectionString\*
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
- System.Data.OleDb.OleDbConnection.ConnectionString
- System.Data.OleDb.OleDbConnection.get\_ConnectionString
- System.Data.OleDb.OleDbConnection.set\_ConnectionString
api\_location:
- System.Data.dll
- System.Data.OleDb.dll
topic\_type:
- apiref
api\_type:
- Assembly
locale: en-us
document\_id: 86257687-34bd-1341-e538-3e9a985e1782
document\_version\_independent\_id: 86257687-34bd-1341-e538-3e9a985e1782
updated\_at: 2026-07-01T18:02:00.0000000Z
original\_content\_git\_url: https://github.com/dotnet/dotnet-api-docs-temp/blob/live/xml/System.Data.OleDb/OleDbConnection.xml
gitcommit: https://github.com/dotnet/dotnet-api-docs-temp/blob/5370c326c6650b054fdaefc78fa9d4d44a5bc53e/xml/System.Data.OleDb/OleDbConnection.xml
git\_commit\_id: 5370c326c6650b054fdaefc78fa9d4d44a5bc53e
default\_moniker: net-11.0-pp
site\_name: Docs
depot\_name: Learn.dotnet-api-docs-temp
page\_type: dotnet
page\_kind: property
ms.assetid: System.Data.OleDb.OleDbConnection.ConnectionString\*
description: 'Gets or sets the string used to open a database. '
toc\_rel: dotnet-api-docs-temp/\_splitted/system.data.oledb/toc.json
search.mshattr.devlang: csharp vb fsharp cpp
asset\_id: api/system.data.oledb.oledbconnection.connectionstring
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
source\_path: xml/System.Data.OleDb/OleDbConnection.xml
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/7696cda6-0510-47f6-8302-71bb5d2e28cf
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/69c76c32-967e-4c65-b89a-74cc527db725
platformId: 01ba5192-23c2-c614-8a18-6d86b74db4a1
---
# OleDbConnection.ConnectionString Property
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
- [OleDbConnection.cs](https://github.com/dotnet/dotnet/blob/3551975be08744f0418857c5bed8ab1545c5dd47/src/runtime/src/libraries/System.Data.OleDb/src/OleDbConnection.cs#L54C17-L54C47)
- Source:
- [OleDbConnection.cs](https://github.com/dotnet/runtime/blob/81cabf2857a01351e5ab578947c7403a5b128ad1/src/libraries/System.Data.OleDb/src/OleDbConnection.cs#L51C17-L51C47)
- Source:
- [OleDbConnection.cs](https://github.com/dotnet/runtime/blob/990ebf52fc408ca45929fd176d2740675a67fab8/src/libraries/System.Data.OleDb/src/OleDbConnection.cs#L51C17-L51C47)
- Source:
- [OleDbConnection.cs](https://github.com/dotnet/dotnet/blob/44525024595742ebe09023abe709df51de65009b/src/runtime/src/libraries/System.Data.OleDb/src/OleDbConnection.cs#L54C17-L54C47)
::: moniker range=" net-10.0-pp net-11.0-pp netframework-1.1 netframework-2.0 netframework-3.0 netframework-3.5 netframework-4.0 netframework-4.5 netframework-4.5.1 netframework-4.5.2 netframework-4.6 netframework-4.6.1 netframework-4.6.2 netframework-4.7 netframework-4.7.1 netframework-4.7.2 netframework-4.8 netframework-4.8.1 netstandard-2.0-pp "
Gets or sets the string used to open a database.
```cpp
public:
virtual property System::String ^ ConnectionString { System::String ^ get(); void set(System::String ^ value); };
```
```cpp
public:
property System::String ^ ConnectionString { System::String ^ get(); void set(System::String ^ value); };
```
```csharp
[System.ComponentModel.SettingsBindable(true)]
public override string ConnectionString { get; set; }
```
```csharp
[System.Data.DataSysDescription("OleDbConnection\_ConnectionString")]
public string ConnectionString { get; set; }
```
```csharp
public override string ConnectionString { get; set; }
```
```fsharp
[]
member this.ConnectionString : string with get, set
```
```fsharp
[]
member this.ConnectionString : string with get, set
```
```fsharp
member this.ConnectionString : string with get, set
```
```vb
Public Overrides Property ConnectionString As String
```
```vb
Public Property ConnectionString As String
```
#### Property Value
[String](system.string)
The OLE DB provider connection string that includes the data source name, and other parameters needed to establish the initial connection. The default value is an empty string.
#### Implements
[ConnectionString](system.data.idbconnection.connectionstring#system-data-idbconnection-connectionstring)
- Attributes
- [SettingsBindableAttribute](system.componentmodel.settingsbindableattribute)[DataSysDescriptionAttribute](system.data.datasysdescriptionattribute)
#### Exceptions
[ArgumentException](system.argumentexception)
An invalid connection string argument has been supplied or a required connection string argument has not been supplied.
## Examples
The following example creates an [OleDbConnection](system.data.oledb.oledbconnection) and sets some of its properties in the connection string.
```csharp
static void OpenConnection(string connectionString)
{
using (OleDbConnection connection = new OleDbConnection(connectionString))
{
try
{
connection.Open();
Console.WriteLine("ServerVersion: {0} \nDataSource: {1}",
connection.ServerVersion, connection.DataSource);
}
catch (Exception ex)
{
Console.WriteLine(ex.Message);
}
// The connection is automatically closed when the
// code exits the using block.
}
}
```
```vb
Public Sub OpenConnection(ByVal connectionString As String)
Using connection As New OleDbConnection(connectionString)
Try
connection.Open()
Console.WriteLine("Server Version: {0} DataSource: {1}", \_
connection.ServerVersion, connection.DataSource)
Catch ex As Exception
Console.WriteLine(ex.Message)
End Try
' The connection is automatically closed when the
' code exits the Using block.
End Using
End Sub
```
## Remarks
The [ConnectionString](system.data.oledb.oledbconnection.connectionstring) is designed to match OLE DB connection string format as closely as possible with the following exceptions:
- The "Provider = `value` " clause is required. However, you cannot use "Provider = MSDASQL" because the .NET Framework Data Provider for OLE DB does not support the OLE DB Provider for ODBC (MSDASQL). To access ODBC data sources, use the [OdbcConnection](system.data.odbc.odbcconnection) object that is in the [System.Data.Odbc](system.data.odbc) namespace.
- Unlike ODBC or ADO, the connection string that is returned is the same as the user-set [ConnectionString](system.data.oledb.oledbconnection.connectionstring), minus security information if `Persist Security Info` is set to `false` (default). The .NET Framework Data Provider for OLE DB does not persist or return the password in a connection string unless you set the `Persist Security Info` keyword to `true` (not recommended). To maintain a high level of security, it is strongly recommended that you use the `Integrated Security` keyword with `Persist Security Info` set to `false`.
You can use the [ConnectionString](system.data.oledb.oledbconnection.connectionstring#system-data-oledb-oledbconnection-connectionstring) property to connect to a variety of data sources. The following example illustrates several possible connection strings.
```txt
"Provider=MSDAORA; Data Source=ORACLE8i7;Persist Security Info=False;Integrated Security=Yes"
"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=c:\bin\LocalAccess40.mdb"
"Provider=SQLOLEDB;Data Source=(local);Integrated Security=SSPI"
```
If the `Data Source` keyword is not specified in the connection string, the provider will try to connect to the local server if one is available.
For more information about connection strings, see [Using Connection String Keywords with SQL Server Native Client](https://go.microsoft.com/fwlink/?LinkId=126696).
The [ConnectionString](system.data.oledb.oledbconnection.connectionstring#system-data-oledb-oledbconnection-connectionstring) property can be set only when the connection is closed. Many of the connection string values have corresponding read-only properties. When the connection string is set, these properties are updated, except when an error is detected. In this case, none of the properties are updated. [OleDbConnection](system.data.oledb.oledbconnection) properties return only those settings that are contained in the [ConnectionString](system.data.oledb.oledbconnection.connectionstring).
Resetting the [ConnectionString](system.data.oledb.oledbconnection.connectionstring) on a closed connection resets all connection string values and related properties. This includes the password. For example, if you set a connection string that includes "Initial Catalog= AdventureWorks", and then reset the connection string to "Provider= SQLOLEDB;Data Source= MySQLServer;IntegratedSecurity=SSPI", the [Database](system.data.oledb.oledbconnection.database#system-data-oledb-oledbconnection-database) property is no longer set to AdventureWorks. (The Initial Catalog value of the connection string corresponds to the `Database` property.)
A preliminary validation of the connection string is performed when the property is set. If values for the `Provider`, `Connect Timeout`, `Persist Security Info`, or `OLE DB Services` are included in the string, these values are checked. When an application calls the [Open](system.data.oledb.oledbconnection.open) method, the connection string is fully validated. If the connection string contains invalid or unsupported properties, a run-time exception, such as [ArgumentException](system.argumentexception), is generated.
Caution
It is possible to supply connection information for an [OleDbConnection](system.data.oledb.oledbconnection) in a Universal Data Link (UDL) file; however you should avoid doing so. UDL files are not encrypted and expose connection string information in clear text. Because a UDL file is an external file-based resource to your application, it cannot be secured using the .NET Framework.
The basic format of a connection string includes a series of keyword/value pairs separated by semicolons. The equal sign (=) connects each keyword and its value. To include values that contain a semicolon, single-quote character, or double-quote character, the value must be enclosed in double quotation marks. If the value contains both a semicolon and a double-quote character, the value can be enclosed in single quotation marks. The single quotation mark is also useful if the value starts with a double-quote character. Conversely, the double quotation mark can be used if the value starts with a single quotation mark. If the value contains both single-quote and double-quote characters, the quotation-mark character used to enclose the value must be doubled every time it occurs within the value.
To include preceding or trailing spaces in the string value, the value must be enclosed in either single quotation marks or double quotation marks. Any leading or trailing spaces around integer, Boolean, or enumerated values are ignored, even if enclosed in quotation marks. However, spaces within a string literal keyword or value are preserved. Single or double quotation marks may be used within a connection string without using delimiters (for example, `Data Source= my'Server` or `Data Source= my"Server`) unless a quotation-mark character is the first or last character in the value.
To include an equal sign (=) in a keyword or value, it must be preceded by another equal sign. For example, in the following hypothetical connection string, the keyword is "key=word" and the value is "value".
```txt
"key==word=value"
```
If a specific keyword in a keyword=value pair occurs multiple times in a connection string, the last occurrence listed is used in the value set.
Keywords are not case sensitive.
Caution
You should use caution when constructing a connection string based on user input, for example, when retrieving user ID and password information from a dialog box and appending it to the connection string. The application should make sure that a user cannot embed additional connection-string parameters in these values, for example, entering a password as "validpassword;database= somedb" in an attempt to attach to a different database. If you use the Extended Properties connection string parameter for OLE DB connections, avoid passing user IDs and passwords because you should avoid storing user IDs and passwords in clear text if you can, and because the default setting of `Persist Security Info= false` does not affect the `Extended Properties` parameter.
## Applies to
## See also
- [Connecting to Data Sources](/en-us/dotnet/framework/data/adonet/connecting-to-a-data-source)
::: moniker-end
