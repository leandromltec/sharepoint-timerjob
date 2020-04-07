using Empresa.setor.TimerJobe.Model;
using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

 /* Controller possui as funções reponsáveis pelo enivo de email */

namespace Empresa.setor.TimerJob.Controller
{
    public class ControllerEmail
    {
        //Envia o email através da Classe Email
        public static void envioEmail(SPWeb web, string email, string assunto, string mensagem)
        {
            Email envioEmail = new Email();

            envioEmail.Para = email;
            envioEmail.Assunto = assunto;
            envioEmail.Corpo = mensagem.ToString();

            envioEmail.EnviaEmail(web);
        }

        //Monta a mensagem do email que será enviando para o responsável
        public static void emailMensagem(string emailResponsavel, SPWeb web, SPListItem solicitacao)
        {
            //Informações do usuário Responsável (campo People Picker)
            SPFieldUserValue responsavelEmail = new SPFieldUserValue(web, solicitacao["ResponsavelSolicitacao"].ToString());

            StringBuilder mensagem = new StringBuilder();
            mensagem.Append("<div style='font-size:15px'><b>Prezado (a)</b>");
            mensagem.Append("<b>Uma Solicitação de Serviço se encontra em atraso</b>");
            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<b>Código da Solicitação: </b>" + solicitacao["Title"].ToString());
            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<b>Responsável pelo Serviço: </b>" + responsavelEmail.User.Name);
            mensagem.Append("</br>");

            mensagem.Append("</br>");

            //Converte para DateTime para obter apenas a data do campo Prazo de Atendimento
            DateTime dataPRazoAtendimento = DateTime.Parse(solicitacao["PrazoFimAtendimento"].ToString());

            mensagem.Append("<b>Prazo para Atendimento:</b> " + dataPRazoAtendimento.Day + "/" + dataPRazoAtendimento.Month + "/" + dataPRazoAtendimento.Year);
            mensagem.Append("</br>");

            mensagem.Append("</br>");
            mensagem.Append("<a href=" + web.Url + "/Lists/SolicitacaoDeServico/DispForm.aspx?ID=" + solicitacao.ID + ">Clique no link para acessar a solicitação</a></div>");

            envioEmail(web, emailResponsavel, "Solicitação em atraso - Código: " + solicitacao["Title"].ToString(), mensagem.ToString());
           

        }

    }
}
