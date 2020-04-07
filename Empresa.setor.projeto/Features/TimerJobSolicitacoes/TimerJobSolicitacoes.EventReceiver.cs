using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Empresa.setor.TimerJob.TimerJob;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

 /* Feature responsável pela ativação/desativação do TimerJob */

namespace Empresa.setor.TimerJobe.Features.TimerJobSolicitacoes
{
    
    [Guid("045c5af5-cf5d-41d6-9639-80d2832308c6")]
    public class TimerJobSolicitacoesEventReceiver : SPFeatureReceiver
    {

        const string JobName = "TimerJob Solicitações de Serviço";

        //Cria o TimerJbo no momento que a Feature é ativada
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            try
            {
                //Eleva os privilégios de permissão para criação do TimerJob
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;

                    //Verifica se o Timer Job existe, caso sim, deleta
                    DeleteExistingJob(JobName, parentWebApp);

                    //Cria o TimerJob
                    CreateJob(parentWebApp);
                });
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        //No momento que a Feature for desativada, deleta o TimerJob
        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {

            lock (this)
            {
                try
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate ()
                    {
                        SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;
                        DeleteExistingJob(JobName, parentWebApp);
                    });
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        //Função cria o TimerJob
        private bool CreateJob(SPWebApplication site)
        {
            bool jobCreated = false;
            try
            {
                Empresa.setor.TimerJob.TimerJob.TimerJobSolicitacoes job = new Empresa.setor.TimerJob.TimerJob.TimerJobSolicitacoes(JobName, site);

                //A cada hora exe
                SPHourlySchedule schedule = new SPHourlySchedule();
                schedule.BeginMinute = 1;
                schedule.EndMinute = 5;

                job.Schedule = schedule;

                job.Update();
            }
            catch (Exception)
            {
                return jobCreated;
            }
            return jobCreated;
        }

        //Função deleta o TimerJob caso ele exista
        public bool DeleteExistingJob(string jobName, SPWebApplication site)
        {
            bool jobDeleted = false;
            try
            {
                foreach (SPJobDefinition job in site.JobDefinitions)
                {
                    if (job.Name == jobName)
                    {
                        job.Delete();
                        jobDeleted = true;
                    }
                }
            }
            catch (Exception)
            {
                return jobDeleted;
            }
            return jobDeleted;
        }
    }
}
