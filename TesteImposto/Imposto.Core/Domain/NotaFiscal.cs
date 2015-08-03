using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Xml.Serialization;
using System.Data;

namespace Imposto.Core.Domain
{
    public class NotaFiscal
    {
        public int Id { get; set; }
        public int NumeroNotaFiscal { get; set; }
        public int Serie { get; set; }
        public string NomeCliente { get; set; }

        public string EstadoDestino { get; set; }
        public string EstadoOrigem { get; set; }

        [XmlIgnore]
        public IEnumerable<NotaFiscalItem> ItensDaNotaFiscal { get; set; }

        public NotaFiscal()
        {
            ItensDaNotaFiscal = new List<NotaFiscalItem>();
        }

        public void EmitirNotaFiscal(Pedido pedido)
        {
            this.NumeroNotaFiscal = 99999;
            this.Serie = new Random().Next(Int32.MaxValue);
            this.NomeCliente = pedido.NomeCliente;

            // correção do erro: variáveis invertidas
            this.EstadoDestino = pedido.EstadoDestino;
            this.EstadoOrigem = pedido.EstadoOrigem;
            
            // Conexao ao banco de dados
            String connStr = ConfigurationManager.ConnectionStrings["MyDBConnectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connStr);
            // insertando os valores P_NOTA_FISCAL
            SqlCommand cmd = new SqlCommand("P_NOTA_FISCAL", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@pId", SqlDbType.Int).Value = this.Id;
            cmd.Parameters.Add("@pNumeroNotaFiscal", SqlDbType.Int).Value = this.NumeroNotaFiscal;
            cmd.Parameters.Add("@pSerie", SqlDbType.Int).Value = this.Serie;
            cmd.Parameters.Add("@pNomeCliente", SqlDbType.VarChar).Value = this.NomeCliente;
            cmd.Parameters.Add("@pEstadoDestino", SqlDbType.VarChar).Value = this.EstadoDestino;
            cmd.Parameters.Add("@pEstadoOrigem", SqlDbType.VarChar).Value = this.EstadoOrigem;

            conn.Open();
            cmd.ExecuteNonQuery();

            SqlCommand cmdId = new SqlCommand();

            // obter o max Id gerado para associar aos itens de pedido
            cmdId.CommandText = "Select max(id) from notaFiscal";
            cmdId.CommandType = CommandType.Text;
            cmdId.Connection = conn;

            this.Id = (int)cmdId.ExecuteScalar();

            foreach (PedidoItem itemPedido in pedido.ItensDoPedido)
            {
                NotaFiscalItem notaFiscalItem = new NotaFiscalItem();

                notaFiscalItem.IdNotaFiscal = this.Id;

                // melhora da complexidade ciclomática 
                if (this.EstadoOrigem == "SP")
                {
                    switch (this.EstadoDestino)
                    {
                        case "RJ":
                            notaFiscalItem.Cfop = "6.000";
                            break;
                        case "PE":
                            notaFiscalItem.Cfop = "6.001";
                            break;
                        case "MG":
                            notaFiscalItem.Cfop = "6.002";
                            break;
                        case "PB":
                            notaFiscalItem.Cfop = "6.003";
                            break;
                        case "PR":
                            notaFiscalItem.Cfop = "6.004";
                            break;
                        case "PI":
                            notaFiscalItem.Cfop = "6.005";
                            break;
                        case "RO":
                            notaFiscalItem.Cfop = "6.006";
                            break;
                        case "SE":
                            notaFiscalItem.Cfop = "6.007";
                            break;
                        case "TO":
                            notaFiscalItem.Cfop = "6.008";
                            break;
                        case "PA":
                            notaFiscalItem.Cfop = "6.010";
                            break;
                    }
                }
                else if (this.EstadoOrigem == "MG")
                {
                    switch (this.EstadoDestino)
                    {
                        case "RJ":
                            notaFiscalItem.Cfop = "6.000";
                            break;
                        case "PE":
                            notaFiscalItem.Cfop = "6.001";
                            break;
                        case "MG":
                            notaFiscalItem.Cfop = "6.002";
                            break;
                        case "PB":
                            notaFiscalItem.Cfop = "6.003";
                            break;
                        case "PR":
                            notaFiscalItem.Cfop = "6.004";
                            break;
                        case "PI":
                            notaFiscalItem.Cfop = "6.005";
                            break;
                        case "RO":
                            notaFiscalItem.Cfop = "6.006";
                            break;
                        case "SE":
                            notaFiscalItem.Cfop = "6.007";
                            break;
                        case "TO":
                            notaFiscalItem.Cfop = "6.008";
                            break;
                        case "PA":
                            notaFiscalItem.Cfop = "6.010";
                            break;
                    }
                }

                if (this.EstadoDestino == this.EstadoOrigem)
                {
                    notaFiscalItem.TipoIcms = "60";
                    notaFiscalItem.AliquotaIcms = 0.18;
                }
                else
                {
                    notaFiscalItem.TipoIcms = "10";
                    notaFiscalItem.AliquotaIcms = 0.17;
                }

                if (notaFiscalItem.Cfop == "6.009")
                {
                    notaFiscalItem.BaseIcms = itemPedido.ValorItemPedido * 0.90; //redução de base
                }
                else
                {
                    notaFiscalItem.BaseIcms = itemPedido.ValorItemPedido;
                }
                notaFiscalItem.ValorIcms = notaFiscalItem.BaseIcms * notaFiscalItem.AliquotaIcms;

                if (itemPedido.Brinde)
                {
                    notaFiscalItem.TipoIcms = "60";
                    notaFiscalItem.AliquotaIcms = 0.18;
                    notaFiscalItem.ValorIcms = notaFiscalItem.BaseIcms * notaFiscalItem.AliquotaIcms;
                }

                notaFiscalItem.NomeProduto = itemPedido.NomeProduto;
                notaFiscalItem.CodigoProduto = itemPedido.CodigoProduto;

                /******* Adiciono valores de IPI ******/
                //Valor Base Ipi
                notaFiscalItem.BaseIpi = itemPedido.ValorItemPedido;
                //Aliquota Ipi
                if (itemPedido.Brinde)
                {
                    // Se for Brinde
                    notaFiscalItem.AliquotaIpi = 0;
                }
                else
                {
                    // Se não é Brinde
                    notaFiscalItem.AliquotaIpi = 0.10;
                }
                //Valor Ipi = Base Cálculo Ipi * Aliquota Ipi
                notaFiscalItem.ValorIpi = notaFiscalItem.BaseIpi * notaFiscalItem.AliquotaIpi;

                // Valor Desconto para clientes onde EstadoDestino seja Sudoeste.
                if (this.EstadoDestino == "SP" || this.EstadoDestino == "RJ" || this.EstadoDestino == "ES" || this.EstadoDestino == "MG")
                {
                    notaFiscalItem.Desconto = 0.10;
                }

                // insertando os valores P_NOTA_FISCAL_ITEM
                SqlCommand cmdIt = new SqlCommand("P_NOTA_FISCAL_ITEM", conn);
                cmdIt.CommandType = CommandType.StoredProcedure;
                cmdIt.Parameters.Add("@pId", SqlDbType.Int).Value = notaFiscalItem.Id;
                cmdIt.Parameters.Add("@pIdNotaFiscal", SqlDbType.Int).Value = notaFiscalItem.IdNotaFiscal;
                cmdIt.Parameters.Add("@pCfop", SqlDbType.VarChar).Value = notaFiscalItem.Cfop;
                cmdIt.Parameters.Add("@pTipoIcms", SqlDbType.VarChar).Value = notaFiscalItem.TipoIcms;
                cmdIt.Parameters.Add("@pBaseIcms", SqlDbType.Decimal).Value = notaFiscalItem.BaseIcms;
                cmdIt.Parameters.Add("@pAliquotaIcms", SqlDbType.Decimal).Value = notaFiscalItem.AliquotaIcms;
                cmdIt.Parameters.Add("@pValorIcms", SqlDbType.Decimal).Value = notaFiscalItem.ValorIcms;
                cmdIt.Parameters.Add("@pNomeProduto", SqlDbType.VarChar).Value = notaFiscalItem.NomeProduto;
                cmdIt.Parameters.Add("@pCodigoProduto", SqlDbType.VarChar).Value = notaFiscalItem.CodigoProduto;
                cmdIt.Parameters.Add("@pBaseIpi", SqlDbType.Decimal).Value = notaFiscalItem.BaseIpi;
                cmdIt.Parameters.Add("@pAliquotaIpi", SqlDbType.Decimal).Value = notaFiscalItem.AliquotaIpi;
                cmdIt.Parameters.Add("@pValorIpi", SqlDbType.Decimal).Value = notaFiscalItem.ValorIpi;
                cmdIt.Parameters.Add("@pDesconto", SqlDbType.Decimal).Value = notaFiscalItem.Desconto;

                cmdIt.ExecuteNonQuery();
            }

            // Persistência de dados XML
            System.Xml.Serialization.XmlSerializer writer =
                new System.Xml.Serialization.XmlSerializer(typeof(NotaFiscal));

            var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "//NotaFiscal_" + this.Id +".xml";
            System.IO.FileStream file = System.IO.File.Create(path);

            writer.Serialize(file, this);
            file.Close();

            // fechando db
            conn.Close();
        }
    }
}
