//------------------------------------------------------------------------------
// <auto-generated>
//    Этот код был создан из шаблона.
//
//    Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//    Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Template4333
{
    using System;
    using System.Collections.Generic;
    using System.Text.Json;
    using System.Text.Json.Serialization;

    public partial class Ord
    {
        public int Id { get; set; }

        [JsonPropertyName("CodeOrder")]
        public string Id_order { get; set; }

        [JsonPropertyName("CreateDate")]
        public string Date_of_creation { get; set; }

        [JsonPropertyName("CreateTime")]
        public string Creation_time { get; set; }

        [JsonPropertyName("CodeClient")]
        public string Id_client { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }

        [JsonPropertyName("ClosedDate")]
        public string Closing_date { get; set; }

        [JsonPropertyName("ProkatTime")]
        public string Rental_time { get; set; }
    }
}
