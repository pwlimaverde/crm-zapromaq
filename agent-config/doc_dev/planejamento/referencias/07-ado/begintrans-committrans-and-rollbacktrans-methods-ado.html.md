<!-- fonte: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado?view=sql-server-ver15 | obtido em: 2026-09-19 | convertido de HTML -->

---
layout: Conceptual
monikers:
- sql-server-ver15
defaultMoniker: sql-server-ver15
versioningType: Ranged
title: BeginTrans, CommitTrans, and RollbackTrans Methods (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
canonicalUrl: https://learn.microsoft.com/en-us/previous-versions/sql/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado?view=sql-server-ver15
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
description: BeginTrans, CommitTrans, and RollbackTrans Methods (ADO)
author: rothja
ms.date: 2017-01-19T00:00:00.0000000Z
ms.service: sql
ms.subservice: ado
ms.topic: reference
apitype: COM
locale: en-us
document\_id: 020900d2-6891-6413-11c1-daf26f21ead2
document\_version\_independent\_id: e5cfa7f3-b99e-7a1e-b9f7-e7110c026266
updated\_at: 2025-01-30T21:32:00.0000000Z
original\_content\_git\_url: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/live/docs-archive-a/sql-server-ver15/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado.md
gitcommit: https://github.com/MicrosoftDocs/sql-docs-archive-pr/blob/5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3/docs-archive-a/sql-server-ver15/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado.md
git\_commit\_id: 5978cb55ab66f4c12d9c3ffbc395c7204f7d38e3
default\_moniker: sql-server-ver15
site\_name: Docs
depot\_name: MSDN.sql-archive-docset-a
page\_type: conceptual
toc\_rel: ../../../connect/toc.json
feedback\_product\_url: ''
word\_count: 472
asset\_id: ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado
moniker\_range\_name: d14ef2ce2d004fb49f8b613ef22df99e
monikers:
- sql-server-ver15
item\_type: Content
source\_path: docs-archive-a/sql-server-ver15/ado/reference/ado-api/begintrans-committrans-and-rollbacktrans-methods-ado.md
platformId: b90d80b3-58c3-8de6-d7fc-8eab627e1f17
---
# BeginTrans, CommitTrans, and RollbackTrans Methods (ADO) - ActiveX Data Objects (ADO) | Microsoft Learn
These transaction methods manage transaction processing within a [Connection](connection-object-ado) object as follows:
- \*\*BeginTrans\*\* Begins a new transaction.
- \*\*CommitTrans\*\* Saves any changes and ends the current transaction. It may also start a new transaction.
- \*\*RollbackTrans\*\* Cancels any changes made during the current transaction and ends the transaction. It may also start a new transaction.
## Syntax
```
level = object.BeginTrans()
object.BeginTrans
object.CommitTrans
object.RollbackTrans
```
## Return Value
\*\*BeginTrans\*\* can be called as a function that returns a \*\*Long\*\* variable indicating the nesting level of the transaction.
#### Parameters
\*object\* A \*\*Connection\*\* object.
## Connection
Use these methods with a \*\*Connection\*\* object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.
Note
Not all providers support transactions. Verify that the provider-defined property "\*\*Transaction DDL\*\*" appears in the \*\*Connection\*\* object's [Properties](properties-collection-ado) collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.
After you call the \*\*BeginTrans\*\* method, the provider will no longer instantaneously commit changes you make until you call \*\*CommitTrans\*\* or \*\*RollbackTrans\*\* to end the transaction.
For providers that support nested transactions, calling the \*\*BeginTrans\*\* method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling \*\*CommitTrans\*\* or \*\*RollbackTrans\*\* affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.
Calling the \*\*CommitTrans\*\* method saves changes made within an open transaction on the connection and ends the transaction. Calling the \*\*RollbackTrans\*\* method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.
Depending on the \*\*Connection\*\* object's [Attributes](attributes-property-ado) property, calling either the \*\*CommitTrans\*\* or \*\*RollbackTrans\*\* methods may automatically start a new transaction. If the \*\*Attributes\*\* property is set to \*\*adXactCommitRetaining\*\*, the provider automatically starts a new transaction after a \*\*CommitTrans\*\* call. If the \*\*Attributes\*\* property is set to \*\*adXactAbortRetaining\*\*, the provider automatically starts a new transaction after a \*\*RollbackTrans\*\* call.
## Remote Data Service
The \*\*BeginTrans\*\*, \*\*CommitTrans\*\*, and \*\*RollbackTrans\*\* methods are not available on a client-side \*\*Connection\*\* object.
## Applies To
[Connection Object (ADO)](connection-object-ado)
