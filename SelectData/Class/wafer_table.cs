using FreeSql.DatabaseModel;using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json;
using FreeSql.DataAnnotations;

namespace SelectData
{

	[JsonObject(MemberSerialization.OptIn), Table(DisableSyncStructure = true)]
	public partial class wafer_table {

		/// <summary>
		/// 自增主键
		/// </summary>
		[JsonProperty, Column(IsPrimary = true, IsIdentity = true)]
		public int id { get; set; }

		/// <summary>
		/// 芯片编号
		/// </summary>
		[JsonProperty]
		public string chip_code { get; set; }

		/// <summary>
		/// 自动创建时间
		/// </summary>
		[JsonProperty, Column(DbType = "timestamp", InsertValueSql = "CURRENT_TIMESTAMP")]
		public DateTime? create_time { get; set; }

		/// <summary>
		/// 自动修改时间
		/// </summary>
		[JsonProperty, Column(DbType = "timestamp", InsertValueSql = "CURRENT_TIMESTAMP")]
		public DateTime? update_time { get; set; }

		/// <summary>
		/// 晶圆编号
		/// </summary>
		[JsonProperty]
		public string wafer_code { get; set; }

	}

}
