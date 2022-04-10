using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace code_order_to_sap
{
    internal class Program
    {
        /// <summary>
        /// Main Test Method
        /// </summary>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {
            /* Declare the company variable - the connection */
            Company company = null;

            Console.WriteLine("Connecting to SAP...");

            /* Connect returns if connection has been established as bool */
            var isConnected = Connect(ref company);

            Console.WriteLine(Environment.NewLine + "Connected; adding order now...");

            var additionResult = AddSalesOrder(company);

            Console.WriteLine(string.Format("{0}Sales Order Addition is Successful = [{1}] and Message is = [{2}]", Environment.NewLine, additionResult.Item1, additionResult.Item2));

            /* Disconnect also released the held memory */
            Disconnect(ref company);

            Console.WriteLine(Environment.NewLine + "Disconnected. Press any key to exit.");
            Console.ReadKey();
        }

        static Tuple<bool, string> AddSalesOrder(Company company)
        {
            /* New Order Instance */
            Documents sapComSalesOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

            /* Header Fields */
            SetOrderHeader(sapComSalesOrder);

            /* Order Lines */
            SetOrderLines(sapComSalesOrder);

            /* Address Setup */
            SetOrderAddresses(sapComSalesOrder);

            /* Add Freight Charges */
            SetDocFreight(sapComSalesOrder);

            var operationMessage = "";

            // Post to SAP
            var isAdditionSuccessful = sapComSalesOrder.Add() == 0;

            if (!isAdditionSuccessful)
            {
                /* Get SAP error in case of unsuccessful addition */
                operationMessage = company.GetLastErrorDescription();
            }

            return new Tuple<bool, string>(isAdditionSuccessful, operationMessage);
        }

        /// <summary>
        /// Sets the sales order header fields
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetOrderHeader(Documents sapComSalesOrder)
        {
            sapComSalesOrder.CardCode = "C20000";
            sapComSalesOrder.DocDate = DateTime.Now;
            sapComSalesOrder.DocDueDate = DateTime.Now.AddDays(1);
            sapComSalesOrder.DocTotal = 34;
            sapComSalesOrder.DocCurrency = "$";
        }

        /// <summary>
        /// Sets order lines
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set them on</param>
        static void SetOrderLines(Documents sapComSalesOrder)
        {
            /* Setting two lines */
            var sapComSalesOrderLine = sapComSalesOrder.Lines;
            SetOrderLastLine(sapComSalesOrderLine, "A00001", "Different Description than SAP's", 2, 10, "$", "01", "CA");
            sapComSalesOrderLine.Add();
            SetOrderLastLine(sapComSalesOrderLine, "A00002", "Different Description than SAP's 2", 1, 5, "$", "01", "CA");
        }

        /// <summary>
        /// Sets the last line of the sales order according to the provided parameters
        /// </summary>
        /// <param name="sapComSalesOrderLine">The SAP order line instance to set it on</param>
        /// <param name="itemCode"></param>
        /// <param name="description"></param>
        /// <param name="quantity"></param>
        /// <param name="unitPrice"></param>
        /// <param name="currency"></param>
        /// <param name="warehouseCode"></param>
        /// <param name="vatCode"></param>
        static void SetOrderLastLine(Document_Lines sapComSalesOrderLine, string itemCode, string description, double quantity, double unitPrice, string currency, string warehouseCode, string vatCode)
        {
            sapComSalesOrderLine.SetCurrentLine(sapComSalesOrderLine.Count - 1);
            sapComSalesOrderLine.ItemCode = itemCode;
            sapComSalesOrderLine.ItemDescription = description;            
            sapComSalesOrderLine.Quantity = quantity;
            sapComSalesOrderLine.UnitPrice = unitPrice;
            sapComSalesOrderLine.Currency = currency;
            sapComSalesOrderLine.WarehouseCode = warehouseCode;
            sapComSalesOrderLine.VatGroup = vatCode;
        }

        /// <summary>
        /// Sets the sales order addresses
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetOrderAddresses(Documents sapComSalesOrder)
        {
            /* Addresses */
            sapComSalesOrder.AddressExtension.BillToStreet = "Billing";
            sapComSalesOrder.AddressExtension.BillToStreetNo = "Clockhouse Place";
            sapComSalesOrder.AddressExtension.BillToBuilding = "Bedfond Road";
            sapComSalesOrder.AddressExtension.BillToCity = "Feltham";
            sapComSalesOrder.AddressExtension.BillToCountry = "GB";
            sapComSalesOrder.AddressExtension.BillToZipCode = "TW14 8HD";

            sapComSalesOrder.AddressExtension.ShipToStreet = "Shipping";
            sapComSalesOrder.AddressExtension.ShipToStreetNo = "Clockhouse Place";
            sapComSalesOrder.AddressExtension.ShipToBuilding = "Bedfond Road";
            sapComSalesOrder.AddressExtension.ShipToCity = "Feltham";            
            sapComSalesOrder.AddressExtension.ShipToCountry = "GB";
            sapComSalesOrder.AddressExtension.ShipToZipCode = "TW14 8HD";
        }

        /// <summary>
        /// Sets Freight Charges on the SAP Document (the sales order)
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetDocFreight(Documents sapComSalesOrder)
        {
            sapComSalesOrder.Expenses.Remarks = "Manual Remark";
            sapComSalesOrder.Expenses.ExpenseCode = 1;
            sapComSalesOrder.Expenses.VatGroup = "CA";
            sapComSalesOrder.Expenses.TaxCode = "CA";
            sapComSalesOrder.Expenses.LineTotal = 4;
            sapComSalesOrder.Expenses.DistributionMethod = BoAdEpnsDistribMethods.aedm_RowTotal;
        }

        /// <summary>
        /// Connects to the provided company.
        /// </summary>
        /// <param name="company">Provide uninstantiated</param>
        /// <returns>True if connection was extablished and False if connection could not be done</returns>
        static bool Connect(ref Company company)
        {
            if (company == null)
            {
                company = new Company();
            }

            if (!company.Connected)
            {
                /* Server connection details */
                company.Server = "db-server-name";
                company.DbServerType = BoDataServerTypes.dst_MSSQL2016;
                company.DbUserName = "sa";
                company.DbPassword = "server-password";
                company.UseTrusted = false;

                /* SAP connection details: DB, SAP User and SAP Password */
                company.CompanyDB = "sap-company/database";
                company.UserName = "sap-user";
                company.Password = "sap-password";

                /* In case the SAP license server is kept in a different location (in most cases can be left empty) */
                company.LicenseServer = "";

                var isSuccessful = company.Connect() == 0;

                return isSuccessful;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Disconnects and releases the held memory (RAM)
        /// </summary>
        /// <param name="company"></param>
        static void Disconnect(ref Company company)
        {
            if (company != null
                && company.Connected)
            {
                company.Disconnect();

                Release(ref company);
            }
        }

        /// <summary>
        /// Re-usable method for releasing COM-held memory
        /// </summary>
        /// <typeparam name="T">Type of object to be released</typeparam>
        /// <param name="obj">The instance to be released</param>
        static void Release<T>(ref T obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { }
            finally
            {
                obj = default(T);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
