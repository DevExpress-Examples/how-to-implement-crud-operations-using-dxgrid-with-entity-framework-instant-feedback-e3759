using System;
using System.Data;
using System.Data.Objects;
using System.Windows;
using System.Windows.Controls;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;

namespace EntityInstantFeedbackCRUD {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        NorthwindEntities context;
        Control control;
        Customers newCustomer;
        Customers customerToEdit;

        public MainWindow() {
            InitializeComponent();
            context = (NorthwindEntities)FindResource("NorthwindEntities");
        }

        private void Add_Click(object sender, RoutedEventArgs e) {
            newCustomer = CreateNewCustomer();
            EditCustomer(newCustomer, "Add new customer", CloseAddNewCustomerHandler);
        }
        private void Remove_Click(object sender, RoutedEventArgs e) {
            DeleteSelectedCustomer(view.FocusedRowHandle);
        }
        private void Edit_Click(object sender, RoutedEventArgs e) {
            EditSelectedCustomer(view.FocusedRowHandle);
        }
        private void view_RowDoubleClick(object sender, RowDoubleClickEventArgs e) {
            EditSelectedCustomer(e.HitInfo.RowHandle);
        }
        private void view_KeyDown(object sender, System.Windows.Input.KeyEventArgs e) {
            if(e.Key == System.Windows.Input.Key.Delete) {
                DeleteSelectedCustomer(view.FocusedRowHandle);
            }
            if(e.Key == System.Windows.Input.Key.Enter) {
                EditSelectedCustomer(view.FocusedRowHandle);
            }
        }

        string GetCustomerIDByRowHandle(int rowHandle) {
            return (string)grid.GetCellValue(rowHandle, colCustomerID);
        }
        void FindCustomerByIDAndProcess(string customerID, Action<Customers> action) {
            try {
                EntityKey entityKey = new EntityKey("NorthwindEntities.Customers", "CustomerID", customerID);
                Customers customer = (Customers)context.GetObjectByKey(entityKey);
                action(customer);
            } catch(Exception ex) {
                HandleException(ex);
            }
        }

        Customers CreateNewCustomer() {
            Customers newCustomer = new Customers();
            newCustomer.CustomerID = GenerateCustomerID();
            return newCustomer;
        }
        string GenerateCustomerID() {
            const int IDLength = 5;
            string result = String.Empty;
            Random rnd = new Random();
            for(int i = 0; i < IDLength; i++) {
                result += Convert.ToChar(rnd.Next(65, 90));
            }
            return result;
        }
        void DeleteSelectedCustomer(int rowHandle) {
            if(rowHandle < 0) return;
            if(MessageBox.Show("Do you really want to delete the selected customer?", "Delete Customer", MessageBoxButton.OKCancel) != MessageBoxResult.OK) return;
            FindCustomerByIDAndProcess(GetCustomerIDByRowHandle(rowHandle), customer => { context.Customers.DeleteObject(customer); SaveChandes(); });
        }
        void EditSelectedCustomer(int rowHandle) {
            if(rowHandle < 0) return;
            FindCustomerByIDAndProcess(GetCustomerIDByRowHandle(rowHandle), customer => {
                customerToEdit = customer; EditCustomer(customerToEdit, "Edit customer", CloseEditCustomerHandler);
            });
        }
        void EditCustomer(Customers customer, string windowTitle, DialogClosedDelegate closedDelegate) {
            control = new ContentControl { Template = (ControlTemplate)Resources["EditRecordTemplate"] };
            control.DataContext = customer;

            FloatingContainer.ShowDialogContent(control, grid, Size.Empty, new FloatingContainerParameters()
            {
                Title = windowTitle,
                AllowSizing = false,
                ClosedDelegate = closedDelegate
            });
        }
        void CloseAddNewCustomerHandler(bool? close) {
            if(close ?? false) {
                context.Customers.AddObject(newCustomer);
                SaveChandes();
            }
            control = null;
            newCustomer = null;
        }
        void CloseEditCustomerHandler(bool? close) {
            if(close ?? false) {
                SaveChandes();
            }
            control = null;
            customerToEdit = null;
        }

        void SaveChandes() {
            try {
                context.SaveChanges();
            } catch(Exception ex) {
                HandleException(ex);
                DetachFailedEntities();
            }
            entityInstantSource.Refresh();
        }
        void DetachFailedEntities() {
            foreach(ObjectStateEntry stateEntry in context.ObjectStateManager.GetObjectStateEntries(
                EntityState.Added | EntityState.Deleted | EntityState.Modified)) {
                context.Detach(stateEntry.Entity);
            }
        }
        void HandleException(Exception ex) {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK);
        }
    }
}
