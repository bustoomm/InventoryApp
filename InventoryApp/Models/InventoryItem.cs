namespace InventoryApp.Models
{
    public class InventoryItem
    {
        public int Id { get; set; }
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public int  Quantity { get; set; }
        public string Unit { get; set; }
        public string Location { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
