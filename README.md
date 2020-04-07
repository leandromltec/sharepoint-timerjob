# sharepoint-timerjob

O código possui um TimerJob para SharePoint 2013 Server (On-Premisses). Foi desenvolvido para uma empresa de geração de energia com
intuito de informar ao usuário responsável (campo people picker em uma lista) suas solicitações que se encontravam em atraso.
O TimerJob varre a lista de solicitações de serviço, verifica se a data no campo Prazo de Atendimento é menor que a data atual e
busca o email de seu responsável no objeto "user" informando sobre o atraso atráves de uma função utilizando o SMPT do SharePoint. 
Ele é verificado 1 vez a cada hora e analisa se a hora atual é 7 da manhã, caso sim dispara o email para o respectivo usuário informando
alguns dados da solicitação que se encontra em atraso.

No código você encontra:

- Projeto SharePoint 2013 Server criado no Visual Studio (Community)
- Criação de Featue e adição de Event Receiver a nível de Web Application para criação do TimerJob (documento em pdf)
- Linguagem C# aplicada a plataforma .NET Framework 4.5
- MVC básico com a pasta Model organizando os ojetos e Controller realizando chamada de funções
- Envio de mail utilzando o SMTP SharePoint de forma programatica
- Criação de um TimerJob
