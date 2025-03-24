namespace contifico.Models
{
    public class Documento
    {
        public string pos { get; set; }
        public string fecha_emision { get; set; }
        public string tipo_documento { get; set; }
        public string estado { get; set; }
        public string caja_id { get; set; }
        public Cliente cliente { get; set; }
        public string vendedor { get; set; }
        public string descripcion { get; set; }
        public double subtotal_0 { get; set; }
        public double subtotal_12 { get; set; }
        public double iva { get; set; }
        public double total { get; set; }
        public string adicional1 { get; set; }
        public Detalle[] detalles { get; set; }
    }
}
