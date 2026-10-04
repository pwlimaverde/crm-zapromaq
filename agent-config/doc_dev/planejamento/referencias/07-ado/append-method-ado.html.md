<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/append-method-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: Append Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/append-method-ado?view=sql-server-ver15
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
description: Append Method (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 3710fdae-1b1f-fb5d-42f1-95ee17deb2b4
document\_version\_independent\_id: e5a007f4-9157-9108-5f40-832fc687caef
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/append-method-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/append-method-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 842
asset\_id: ado/reference/ado-api/append-method-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/append-method-ado.md
platformId: b780c60f-5fc1-3b49-78a9-fe7d97177e75
---
# Append Method (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
Appends an object to a collection. If the collection is [Fields](fields-collection-ado), a new [Field](field-object) object can be created before it is appended to the collection.
## Syntax
```
collection.Append object
fields.Append Name, Type, DefinedSize, Attrib, FieldValue
```
#### Parameters
\*collection\* A collection object.
\*fields\* A \*\*Fields\*\* collection.
\*object\* An object variable that represents the object to be appended.
\*Name\* A \*\*String\*\* value that contains the name of the new \*\*Field\*\* object, and must not be the same name as any other object in \*fields\*.
\*Type\* A [DataTypeEnum](datatypeenum) value, whose default value is \*\*adEmpty\*\*, that specifies the data type of the new field. The following data types are not supported by ADO, and should not be used when appending new fields to a [Recordset Object (ADO)](recordset-object-ado): \*\*adIDispatch\*\*, \*\*adIUnknown\*\*, \*\*adVariant\*\*.
\*DefinedSize\* Optional. A \*\*Long\*\* value that represents the defined size, in characters or bytes, of the new field. The default value for this parameter is derived from \*Type\*. Fields that have a \*DefinedSize\* greater than 255 bytes are treated as variable length columns. The default for \*DefinedSize\* is unspecified.
\*Attrib\* Optional. A [FieldAttributeEnum](fieldattributeenum) value, whose default value is \*\*adFldDefault\*\*, that specifies attributes for the new field. If this value is not specified, the field will contain attributes derived from \*Type\*.
\*FieldValue\* Optional. A \*\*Variant\*\* that represents the value for the new field. If not specified, the field is appended with a null value.
## Remarks
## Parameters Collection
You must set the [Type](type-property-ado) property of a [Parameter](parameter-object) object before appending it to the [Parameters](parameters-collection-ado) collection. If you select a variable-length data type, you must also set the [Size](size-property-ado-parameter) property to a value greater than zero.
Describing parameters yourself minimizes calls to the provider and therefore improves performance when you use stored procedures or parameterized queries. However, you must know the properties of the parameters associated with the stored procedure or parameterized query that you want to call.
Use the [CreateParameter](createparameter-method-ado) method to create \*\*Parameter\*\* objects with the appropriate property settings and use the \*\*Append\*\* method to add them to the [Parameters](parameters-collection-ado) collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must use this method to manually populate the \*\*Parameters\*\* collection in order to use parameters at all.
## Fields Collection
The \*FieldValue\* parameter is only valid when adding a \*\*Field\*\* object to a [Record](record-object-ado) object, not to a \*\*Recordset\*\* object. With a \*\*Record\*\* object, you can append fields and provide values at the same time. With a \*\*Recordset\*\* object, you must create fields while the \*\*Recordset\*\* is closed, and then open the \*\*Recordset\*\* and assign values to the fields.
Note
For new \*\*Field\*\* objects that have been appended to the \*\*Fields\*\* collection of a \*\*Record\*\* object, the [Value](value-property-ado) property must be set before any other \*\*Field\*\* properties can be specified. First, a specific value for the \*\*Value\*\* property must have been assigned and [Update](update-method) on the \*\*Fields\*\* collection called. Then, other properties such as [Type](type-property-ado) or [Attributes](attributes-property-ado) can be accessed. \*\*Field\*\* objects of the following data types (\*\*DataTypeEnum\*\*) cannot be appended to the \*\*Fields\*\* collection and will cause an error to occur: \*\*adArray\*\*, \*\*adChapter\*\*, \*\*adEmpty\*\*, \*\*adPropVariant\*\*, and \*\*adUserDefined\*\*. Also, the following data types are not supported by ADO: \*\*adIDispatch\*\*, \*\*adIUnknown\*\*, and \*\*adIVariant\*\*. For these types, no error will occur when appended, but usage can produce unpredictable results including memory leaks.
## Recordset
If you do not set the [CursorLocation](cursorlocation-property-ado) property before calling the \*\*Append\*\* method, \*\*CursorLocation\*\* will be set to \*\*adUseClient\*\* (a [CursorLocationEnum](cursorlocationenum) value) automatically when the [Open](open-method-ado-recordset) method of the [Recordset](recordset-object-ado) object is called.
A run-time error will occur if the \*\*Append\*\* method is called on the \*\*Fields\*\* collection of an open \*\*Recordset\*\*, or on a \*\*Recordset\*\* where the [ActiveConnection](activeconnection-property-ado) property has been set. You can only append fields to a \*\*Recordset\*\* that is not open and has not yet been connected to a data source. This is typically the case when a \*\*Recordset\*\* object is fabricated with the [CreateRecordset](../rds-api/createrecordset-method-rds) method or assigned to an object variable.
## Record
A run-time error will not occur if the \*\*Append\*\* method is called on the \*\*Fields\*\* collection of an open \*\*Record\*\*. The new field will be added to the \*\*Fields\*\* collection of the \*\*Record\*\* object. If the \*\*Record\*\* was derived from a \*\*Recordset\*\*, the new field will not appear in the \*\*Fields\*\* collection of the \*\*Recordset\*\* object.
A non-existent field can be created and appended to the \*\*Fields\*\* collection by assigning a value to the field object as if it already existed in the collection. The assignment will trigger the automatic creation and appending of the \*\*Field\*\* object, and then the assignment will be completed.
After appending a \*\*Field\*\* to the \*\*Fields\*\* collection of a \*\*Record\*\* object, call the \*\*Update\*\* method of the \*\*Fields\*\* collection to save the change.
## Applies To
- [Fields Collection (ADO)](fields-collection-ado)
- [Parameters Collection (ADO)](parameters-collection-ado)
