using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace exceltest.Controllers
{
    public static class DataInfo
    {
        public static List<EquipmentRevenueViewModel> Getdata()
        {
            return new List<EquipmentRevenueViewModel>
            {
                new EquipmentRevenueViewModel
                {
                    Today = DateTime.Now,
                    Category = "Category",
                    Equipment = "Equipment",
                    Department = "Department",
                    OrderNo = 14,
                    DaysBilled = 1,
                    DaysRented = 2,
                    EquipDesc = "EquipDesc",
                    OrderDesc = "OrderDesc",
                    QtyBilled = 1,
                    Vendor = "Vendor",
                    QtyRented = 487
                },
                new EquipmentRevenueViewModel
                {
                    Today = DateTime.Now.AddDays(1),
                    Category = "Category2",
                    Equipment = "Equipment2",
                    Department = "Department2",
                    OrderNo = 14,
                    DaysBilled = 1,
                    DaysRented = 2,
                    EquipDesc = "EquipDesc2",
                    OrderDesc = "OrderDesc2",
                    QtyBilled = 1,
                    Vendor = "Vendor2",
                    QtyRented = 120
                }
            };
        }
	}
    public class EquipmentRevenueViewModel
    {
        public DateTime Today { get; set; }
        [StringLength(20)]
        public string Equipment { get; set; }
        [StringLength(20)]
        public string Department { get; set; }
        [StringLength(22)]
        public string Category { get; set; }
        
        [DisplayName("Equipment&Description")]
        [StringLength(50)]
        public string EquipDesc { get; set; }

        [StringLength(10)]
        [DisplayName("OrderNo#")]
        public int? OrderNo { get; set; }

        [StringLength(40)]
        [DisplayName("Order&Description")]
        public string OrderDesc { get; set; }
        public string Vendor { get; set; }
        [DisplayName("Days&Rented")]
        public int? DaysRented { get; set; }
        [DisplayName("Qty&Rented")]
        public int? QtyRented { get; set; }
        [DisplayName("Qty&Billed")]
        public decimal? QtyBilled { get; set; }
        [DisplayName("Days&Billed")]
        public decimal? DaysBilled { get; set; }
    }
}