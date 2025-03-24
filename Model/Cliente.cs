namespace contifico.Models
{
    public class Cliente
    {
        public string ruc { get; set; }
        public string cedula { get; set; }
        public string razon_social { get; set; }
        public string telefonos { get; set; }
        public string direccion { get; set; }
        public string tipo { get; set; }
        public string email { get; set; }
        public bool es_extranjero { get; set; }
    }
}
