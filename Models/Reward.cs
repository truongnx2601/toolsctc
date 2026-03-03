namespace ToolsCTC.Models
{
    public class Reward
    {
        public Guid Id { get; set; }
        public string Icon { get; set; }
        public string Text { get; set; }
        public bool Active { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
