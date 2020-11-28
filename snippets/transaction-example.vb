'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Last Modified Date: 27 Nov, 2020
'Description: Example of how to safely create an entry in Inventor's undo stack
Option Explicit On

Const transactionName As String = "My Transaction" 'This string will show up in Inventor's undo stack

Dim app As Inventor.Application = ThisApplication
Dim doc As Document = ThisDoc.Document

Dim oTransaction As Transaction
Try
	'AFTER STARTING THE TRANSACTION, IT MUST BE ENDED OR ABORTED BEFORE
	'THE SCRIPT THROWS AN UNHANDLED EXCEPTION OR ENDS.  OTHERWISE, INVENTOR CAN CRASH.
	oTransaction = app.TransactionManager.StartTransaction(doc, transactionName)
	
	'Code Here... Any modifications to the document will show up as a single entry in the undo stack.
	
	doc.Update2() 'If you need to update the document, do it before the transaction ends, so the update won't create an additional undo entry.
	
	oTransaction.End 'This creates an entry in the undo stack.
	
Catch ex As Exception 'for catching any unhandled exceptions
	oTransaction.Abort 'The document will be rolled back to the point before the transaction was started.
	Throw 'Re-throw the unhandled exception
End Try