using Empresa.setor.TimerJob.Controller;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

/* Classe para construção do TimerJob */

namespace Empresa.setor.TimerJob.TimerJob
{
    //Classe SPJobDefinition define herança para construtor do TimerJob
    public class TimerJobSolicitacoes : SPJobDefinition
    {
        //Construtor TimerJob
        public TimerJobSolicitacoes()
           : base()
        {
        }

        public TimerJobSolicitacoes(string nomeJob, SPService service)
            : base(nomeJob, service, null, SPJobLockType.None)
        {
            this.Title = "Setor - Solicitacoes de Servico em Atraso";
        }



        public TimerJobSolicitacoes(string nomeJob, SPWebApplication webapp)
            : base(nomeJob, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Setor - Solicitacoes de Servico em Atraso";
        }

        
        //Sobreescreve a função do TimerJob que será executada
        public override void Execute(Guid targetInstanceId)
        {
            //Data atual
            DateTime horario = DateTime.Now;

            //Verifica se a hora atual é igual ás 7 da manhã
            if (horario.Hour == 7)
            {
                //WebApplication referente ao timer Job
                SPWebApplication webApp = this.Parent as SPWebApplication;

                SPSite site = webApp.Sites[0].RootWeb.Site;

                //Site Atendimento que contém a lista de solicitações
                using (SPWeb webAtendimento = site.OpenWeb("Atendimento"))
                {
                    //Lista de Solicitações
                    SPList listaSolicitacoes = webAtendimento.Lists["Solicitações de Serviço"];

                    //Query verifica se a Solicitação está em atendimento e se a data da solicitação é menor que a data atual
                    SPQuery queryAtendimento = new SPQuery();
                    queryAtendimento.Query = string.Format("<Where><And><Eq><FieldRef Name='Status' /><Value Type='Choice'>Em Atendimento</Value></Eq>" +
                                              "<Lt><FieldRef Name='PrazoFimAtendimento' /><Value IncludeTimeValue='FALSE' Type='DateTime'>" + Convert.ToDateTime(DateTime.Now.ToString()).ToString("yyyy-MM-ddThh:mm:ssZ") + "</Value></Lt>" +
                                              "</And></Where>");

                    //Coleção de itens que realizados com a consulta
                    SPListItemCollection colecaoSolicitacoesAtrasadas = listaSolicitacoes.GetItems(queryAtendimento);
                    
                    foreach (SPListItem itemSolicitacao in colecaoSolicitacoesAtrasadas)
                    {
                        //Verifica se a solicitação possui Responsável
                        if (itemSolicitacao["ResponsavelSolicitacao"] != null)
                        {

                            SPFieldUserValue responsavelEmail = new SPFieldUserValue(webAtendimento, itemSolicitacao["ResponsavelSolicitacao"].ToString());

                            if (itemSolicitacao["PrazoFimAtendimento"] != null)
                            {

                                DateTime dataPrazoAtendimento = DateTime.Parse(itemSolicitacao["PrazoFimAtendimento"].ToString());

                                  ControllerEmail.emailMensagem(responsavelEmail.User.Email, webAtendimento, itemSolicitacao);

                                

                            }

                        }


                    }
                }

            }

        }
    }


}
