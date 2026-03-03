using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Npgsql;
using System.Data;
using ToolsCTC.Models;

namespace ToolsCTC.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RewardsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public RewardsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private IDbConnection CreateConnection()
        {
            return new NpgsqlConnection(_configuration.GetConnectionString("DefaultConnection"));
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            try
            {
                using var db = CreateConnection();
                var rewards = await db.QueryAsync<Reward>("SELECT * FROM rewards WHERE active = true ORDER BY created_at");
                return Ok(rewards);
            }
            catch(Exception ex)
            {
                Console.WriteLine("LỖI: " + ex.ToString());
                return StatusCode(500, $"Lỗi server: {ex.Message}");
            }
            
        }

        [HttpPost]
        public async Task<IActionResult> Create([FromBody] CreateDTO reward)
        {
            using var db = CreateConnection();
            var sql = "INSERT INTO rewards (icon, text) VALUES (@Icon, @Text)";
            await db.ExecuteAsync(sql, reward);
            return Ok(reward);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> Update(Guid id, [FromBody] Reward reward)
        {
            using var db = CreateConnection();
            var sql = "UPDATE rewards SET icon = @Icon, text = @Text WHERE id = @Id";
            reward.Id = id;
            await db.ExecuteAsync(sql, reward);
            return Ok(reward);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> SoftDelete(Guid id)
        {
            using var db = CreateConnection();
            var sql = "UPDATE rewards SET active = false WHERE id = @Id";
            await db.ExecuteAsync(sql, new { Id = id });
            return NoContent();
        }
    }
}
