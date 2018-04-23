Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.Objects
Imports System.Windows
Imports System.Windows.Controls
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.Grid

Namespace EntityInstantFeedbackCRUD
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Private context As NorthwindEntities
		Private control As Control
		Private newCustomer As Customers
		Private customerToEdit As Customers

		Public Sub New()
			InitializeComponent()
			context = CType(FindResource("NorthwindEntities"), NorthwindEntities)
		End Sub

		Private Sub Add_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			newCustomer = CreateNewCustomer()
			EditCustomer(newCustomer, "Add new customer", AddressOf CloseAddNewCustomerHandler)
		End Sub
		Private Sub Remove_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			DeleteSelectedCustomer(view.FocusedRowHandle)
		End Sub
		Private Sub Edit_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			EditSelectedCustomer(view.FocusedRowHandle)
		End Sub
		Private Sub view_RowDoubleClick(ByVal sender As Object, ByVal e As RowDoubleClickEventArgs)
			EditSelectedCustomer(e.HitInfo.RowHandle)
		End Sub
		Private Sub view_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs)
			If e.Key = System.Windows.Input.Key.Delete Then
				DeleteSelectedCustomer(view.FocusedRowHandle)
			End If
			If e.Key = System.Windows.Input.Key.Enter Then
				EditSelectedCustomer(view.FocusedRowHandle)
			End If
		End Sub

		Private Function GetCustomerIDByRowHandle(ByVal rowHandle As Integer) As String
			Return CStr(grid.GetCellValue(rowHandle, colCustomerID))
		End Function
		Private Sub FindCustomerByIDAndProcess(ByVal customerID As String, ByVal action As Action(Of Customers))
			Try
				Dim entityKey As New EntityKey("NorthwindEntities.Customers", "CustomerID", customerID)
				Dim customer As Customers = CType(context.GetObjectByKey(entityKey), Customers)
				action(customer)
			Catch ex As Exception
				HandleException(ex)
			End Try
		End Sub

		Private Function CreateNewCustomer() As Customers
			Dim newCustomer As New Customers()
			newCustomer.CustomerID = GenerateCustomerID()
			Return newCustomer
		End Function
		Private Function GenerateCustomerID() As String
			Const IDLength As Integer = 5
			Dim result As String = String.Empty
			Dim rnd As New Random()
			For i As Integer = 0 To IDLength - 1
				result &= Convert.ToChar(rnd.Next(65, 90))
			Next i
			Return result
		End Function
		Private Sub DeleteSelectedCustomer(ByVal rowHandle As Integer)
			If rowHandle < 0 Then
				Return
			End If
			If MessageBox.Show("Do you really want to delete the selected customer?", "Delete Customer", MessageBoxButton.OKCancel) <> MessageBoxResult.OK Then
				Return
			End If
			FindCustomerByIDAndProcess(GetCustomerIDByRowHandle(rowHandle), Function(customer) AnonymousMethod1(customer))
		End Sub
		
		Private Function AnonymousMethod1(ByVal customer As Customers) As Boolean
			context.Customers.DeleteObject(customer)
			SaveChandes()
			Return True
		End Function
		Private Sub EditSelectedCustomer(ByVal rowHandle As Integer)
			If rowHandle < 0 Then
				Return
			End If
			FindCustomerByIDAndProcess(GetCustomerIDByRowHandle(rowHandle), Function(customer) AnonymousMethod2(customer))
		End Sub
		
		Private Function AnonymousMethod2(ByVal customer As Customers) As Boolean
			customerToEdit = customer
			EditCustomer(customerToEdit, "Edit customer", AddressOf CloseEditCustomerHandler)
			Return True
		End Function
		Private Sub EditCustomer(ByVal customer As Customers, ByVal windowTitle As String, ByVal closedDelegate As DialogClosedDelegate)
			control = New ContentControl With {.Template = CType(Resources("EditRecordTemplate"), ControlTemplate)}
			control.DataContext = customer

			FloatingContainer.ShowDialogContent(control, grid, Size.Empty, New FloatingContainerParameters() With {.Title = windowTitle, .AllowSizing = False, .ClosedDelegate = closedDelegate})
		End Sub
		Private Sub CloseAddNewCustomerHandler(ByVal close? As Boolean)
			If If(close, False) Then
				context.Customers.AddObject(newCustomer)
				SaveChandes()
			End If
			control = Nothing
			newCustomer = Nothing
		End Sub
		Private Sub CloseEditCustomerHandler(ByVal close? As Boolean)
			If If(close, False) Then
				SaveChandes()
			End If
			control = Nothing
			customerToEdit = Nothing
		End Sub

		Private Sub SaveChandes()
			Try
				context.SaveChanges()
			Catch ex As Exception
				HandleException(ex)
				DetachFailedEntities()
			End Try
			entityInstantSource.Refresh()
		End Sub
		Private Sub DetachFailedEntities()
			For Each stateEntry As ObjectStateEntry In context.ObjectStateManager.GetObjectStateEntries(EntityState.Added Or EntityState.Deleted Or EntityState.Modified)
				context.Detach(stateEntry.Entity)
			Next stateEntry
		End Sub
		Private Sub HandleException(ByVal ex As Exception)
			MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK)
		End Sub
	End Class
End Namespace
