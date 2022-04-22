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

	[JsonObject(MemberSerialization.OptIn), Table(Name = "c-ave-0.11", DisableSyncStructure = true)]
	public partial class c_ave_0_11 {

		[JsonProperty, Column(IsPrimary = true, IsIdentity = true)]
		public int id { get; set; }

		[JsonProperty, Column(Name = "40CH_AX", StringLength = 50)]
		public string _40CH_AX { get; set; }

		[JsonProperty, Column(Name = "40CH_BD_0.5", StringLength = 50)]
		public string _40CH_BD_0_5 { get; set; }

		[JsonProperty, Column(Name = "40CH_BD_1", StringLength = 50)]
		public string _40CH_BD_1 { get; set; }

		[JsonProperty, Column(Name = "40CH_BD_20", StringLength = 50)]
		public string _40CH_BD_20 { get; set; }

		[JsonProperty, Column(Name = "40CH_BD_3", StringLength = 50)]
		public string _40CH_BD_3 { get; set; }

		[JsonProperty, Column(Name = "40CH_IL", StringLength = 50)]
		public string _40CH_IL { get; set; }

		[JsonProperty, Column(Name = "40CH_NX", StringLength = 50)]
		public string _40CH_NX { get; set; }

		[JsonProperty, Column(Name = "40CH_offset", StringLength = 50)]
		public string _40CH_offset { get; set; }

		[JsonProperty, Column(Name = "40CH_PDL", StringLength = 50)]
		public string _40CH_PDL { get; set; }

		[JsonProperty, Column(Name = "40CH_ripple", StringLength = 50)]
		public string _40CH_ripple { get; set; }

		[JsonProperty, Column(Name = "40CH_TX", StringLength = 50)]
		public string _40CH_TX { get; set; }

		[JsonProperty, Column(Name = "40CH工作波段", StringLength = 50)]
		public string _40CH工作波段 { get; set; }

		[JsonProperty, Column(Name = "40CH工作通道", StringLength = 50)]
		public string _40CH工作通道 { get; set; }

		[JsonProperty, Column(Name = "48CH_AX", StringLength = 50)]
		public string _48CH_AX { get; set; }

		[JsonProperty, Column(Name = "48CH_BD_0.5", StringLength = 50)]
		public string _48CH_BD_0_5 { get; set; }

		[JsonProperty, Column(Name = "48CH_BD_1", StringLength = 50)]
		public string _48CH_BD_1 { get; set; }

		[JsonProperty, Column(Name = "48CH_BD_20", StringLength = 50)]
		public string _48CH_BD_20 { get; set; }

		[JsonProperty, Column(Name = "48CH_BD_3", StringLength = 50)]
		public string _48CH_BD_3 { get; set; }

		[JsonProperty, Column(Name = "48CH_IL", StringLength = 50)]
		public string _48CH_IL { get; set; }

		[JsonProperty, Column(Name = "48CH_NX", StringLength = 50)]
		public string _48CH_NX { get; set; }

		[JsonProperty, Column(Name = "48CH_offset", StringLength = 50)]
		public string _48CH_offset { get; set; }

		[JsonProperty, Column(Name = "48CH_PDL", StringLength = 50)]
		public string _48CH_PDL { get; set; }

		[JsonProperty, Column(Name = "48CH_ripple", StringLength = 50)]
		public string _48CH_ripple { get; set; }

		[JsonProperty, Column(Name = "48CH_TX", StringLength = 50)]
		public string _48CH_TX { get; set; }

		[JsonProperty, Column(Name = "48CH工作波段")]
		public string _48CH工作波段 { get; set; }

		[JsonProperty, Column(Name = "48CH工作通道")]
		public string _48CH工作通道 { get; set; }

		/// <summary>
		/// 芯片编号
		/// </summary>
		[JsonProperty]
		public string chip_code { get; set; }

		[JsonProperty, Column(DbType = "timestamp", InsertValueSql = "CURRENT_TIMESTAMP")]
		public DateTime? creat_time { get; set; }

		[JsonProperty, Column(DbType = "timestamp", InsertValueSql = "CURRENT_TIMESTAMP")]
		public DateTime? update_time { get; set; }

		/// <summary>
		/// 晶圆编号
		/// </summary>
		[JsonProperty]
		public string wafer_code { get; set; }

	}

}
