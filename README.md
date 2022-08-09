<h1> Enviando email com python</h1> 

<h2>Iniciando o projeto</h2>

O primeiro passo é instalar a biblioteca para utilizar no projeto, abra o seu terminal e digite o seguinte comando:

![image](https://user-images.githubusercontent.com/77951123/183555434-78bb829e-99a9-4bd3-a88f-381674af04f1.png)

Depois vamos criar uma pasta e dentro dela um arquivo .py com o nome da sua preferência; (use o editor de código que você mais gosta =D)

![image](https://user-images.githubusercontent.com/77951123/183555844-b297a356-2eb1-4325-8449-995f8e2fbcd2.png)

<h2>importanto a Biblioteca</h2>

Dentro do arquivo vamos começar importando a biblioteca pywin32 que te permite automatizar uma série de coisas no Windows, então isso pode te facilitar bastante caso utilize o Windows.

![image](https://user-images.githubusercontent.com/77951123/183556896-8713af4e-b29f-4955-97f7-7d73be34838e.png)


<h2>Definindo os objetos</h2>

Apenas precisamos definir que objeto vamos utilizar que no caso é o OUTLOOK e logo depois precisamos definir uma váriavel vazia:
![image](https://user-images.githubusercontent.com/77951123/183556782-fed02139-c099-4b65-9b6c-814012318b29.png)

<h2>Lidando com várias contas dentro do outlook</h2>
Caso tenha várias contas logadas no seu outlook mas você quer apenas utilizar uma para enviar seus email precisamos criar uma estrutura de FOR: 

![image](https://user-images.githubusercontent.com/77951123/183557633-20f9f0f5-1c53-4b07-bdbb-dbcfc883e579.png)

<h2>E agora é só criar o objeto e definir um IF</h2>

![image](https://user-images.githubusercontent.com/77951123/183557804-ec8c192f-5c77-4548-af65-433cbe2a382c.png)


<H2>Explicando cada campo para edição</h2>
Depois de ter definido o objeto e o IF agora só devemos preencher os seguintes itens;

<b>mail.to<b> Aqui deve-se colocar a sua lista de email para envio;
  ![image](https://user-images.githubusercontent.com/77951123/183558068-d3144bea-380d-4861-aa30-16932ecfbf6f.png)

  <hr>
 
<b>mail.Subject<b> é onde devemos colocar o assunto;
  
 ![image](https://user-images.githubusercontent.com/77951123/183558317-c4b3fa73-b2a2-4f5c-85f7-e129d8b5b394.png)
  
   <hr>
  
<b>mail.CC e mail.BCC<b> Envia uma cópia do email e BCC envia uma cópia oculta do mesmo;
  
  ![image](https://user-images.githubusercontent.com/77951123/183558669-5a4b2311-4ea4-4631-9b58-ec8338dcf6f1.png)
  
  
 <hr>
  
  <b>mailHTMLBody<b> É onde a mágica acontece aqui definimos a mensagem que vamos enviar e a melhor forma de fazer isso é utilizando o HTML como no exemplo abaixo: 
    ![image](https://user-images.githubusercontent.com/77951123/183558888-9095567a-4587-44ef-887a-b78b69eed5aa.png)

    Mas relaxa vc pode usar o <b>mail.Body<b> caso ainda não saiba muito sobre HTML 
    
    
    <h2>Enviando Anexo</h2>
    Para enviar anexo precisamos apenas definir uma váriavel com o caminho do arquivo;
    ![image](https://user-images.githubusercontent.com/77951123/183559062-9b40d121-4036-4f95-b3a8-734b14d3bd7f.png)

    
    
    <h2>E enviar</h2>
    
    E para finalizar;
    
    ![image](https://user-images.githubusercontent.com/77951123/183559150-1fdc3612-bea1-4f3c-bb23-906e2b4145bf.png)
    
    
    <hr>
    
    <h2>Obrigado por ter chegado até aqui! Caso eu tenha te ajudado de alguma forma me siga =D Qualquer dúvida só enviar pelos meios de comunicação que desponibilizo no meu perfil</h2>

    
    

  
  
  










