---
title: Utilisation d’OAuth dans un complément Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 3fac2dd0ca6231684b0b91db80f969787822cf5f
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270697"
---
# <a name="use-the-oauth-authorization-framework-in-an-office-add-in"></a><span data-ttu-id="948f0-102">Utilisation d’OAuth dans un complément Office</span><span class="sxs-lookup"><span data-stu-id="948f0-102">Use the OAuth authorization framework in an Office Add-in</span></span>

<span data-ttu-id="948f0-p101">OAuth est le standard d’authentification ouvert utilisé par les fournisseurs de services en ligne (Office 365, Facebook, Google, SalesForce, LinkedIn, etc.) pour procéder à l’authentification des utilisateurs. OAuth est le protocole d’authentification par défaut pour Azure et Office 365. OAuth est utilisé par les entreprises et les consommateurs.</span><span class="sxs-lookup"><span data-stu-id="948f0-p101">OAuth is the open standard for authorization that online service providers such as Office 365, Facebook, Google, SalesForce, LinkedIn and others use to perform user authentication. The OAuth authorization framework is the default authorization protocol used in Azure and Office 365. The OAuth authorization framework is used in both enterprise (corporate) and consumer scenarios.</span></span>

<span data-ttu-id="948f0-106">Les fournisseurs de services en ligne peuvent proposer des API publiques qui se connectent avec REST.</span><span class="sxs-lookup"><span data-stu-id="948f0-106">Online service providers may provide public APIs exposed via REST.</span></span> <span data-ttu-id="948f0-107">Les développeurs peuvent utiliser ces API publiques dans leurs compléments Office pour lire ou écrire des données au fournisseur de services en ligne.</span><span class="sxs-lookup"><span data-stu-id="948f0-107">Developers can use these public APIs in their Office Add-ins to read or write data to the online service provider.</span></span> <span data-ttu-id="948f0-108">Intégrer les données des fournisseurs de services en ligne dans un complément augmente sa valeur et entraîne de meilleurs taux d’adoption.</span><span class="sxs-lookup"><span data-stu-id="948f0-108">Integrating data from online service providers in an add-in increases its value, which leads to greater user adoption.</span></span> <span data-ttu-id="948f0-109">Lorsque vous utilisez ces API dans votre complément, les utilisateurs doivent s’authentifier avec OAuth.</span><span class="sxs-lookup"><span data-stu-id="948f0-109">When using these APIs in your add-in, users will be required to authenticate using the OAuth authorization framework.</span></span>

<span data-ttu-id="948f0-p103">Cette rubrique explique comment mettre en œuvre un flux d’authentification dans votre complément pour procéder à l’authentification des utilisateurs. Les segments de code inclus dans cette rubrique sont extraits de l’exemple de code [Office-Add-in-NodeJS-ServerAuth](https://github.com/OfficeDev/Office-Add-in-NodeJS-ServerAuth).</span><span class="sxs-lookup"><span data-stu-id="948f0-p103">This topic describes how to implement an authentication flow in your add-in to perform user authentication. Code segments included in this topic are taken from the [Office-Add-in-NodeJS-ServerAuth](https://github.com/OfficeDev/Office-Add-in-NodeJS-ServerAuth) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="948f0-p104">Pour des raisons de sécurité, les navigateurs ne sont pas autorisés à afficher les pages de connexion dans un IFrame. Selon la version d’Office utilisée par vos clients, notamment les versions web, votre complément s’affiche dans un IFrame. C’est pourquoi il est indispensable de savoir comment gérer le flux d’authentification. </span><span class="sxs-lookup"><span data-stu-id="948f0-p104">For security reasons, browsers are not allowed to display sign-in pages in an IFrame. Depending on the version of Office that your customers use, most notably web-based versions, your add-in is displayed in an IFrame. This imposes some considerations on how to manage the authentication flow.</span></span> 

<span data-ttu-id="948f0-115">Le schéma suivant présente les composants nécessaires, ainsi que le flux d’événements qui se produisent lors de l’implémentation de l’authentification dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="948f0-115">The following diagram shows the required components and the flow of events that occur when implementing authentication in your add-in.</span></span>

![Réalisation d’une authentification OAuth dans un complément Office](../images/oauth-in-office-add-in.png)

<span data-ttu-id="948f0-117">Le schéma illustre la façon dont les composants requis suivants sont utilisés :</span><span class="sxs-lookup"><span data-stu-id="948f0-117">The diagram shows how the following required components are used:</span></span>


- <span data-ttu-id="948f0-p105">Office exécute un complément de volet Office sur l’ordinateur de l’utilisateur. Votre complément ouvre une fenêtre contextuelle pour démarrer le flux d’authentification. Les compléments ne peuvent pas démarrer les flux d’authentification directement, car ils sont exécutés dans un IFRAME (selon la plateforme utilisée). Pour des raisons de sécurité, les pages de connexion OAuth ne peuvent pas être affichées dans un IFRAME.</span><span class="sxs-lookup"><span data-stu-id="948f0-p105">Office runs a task pane add-in on the user's computer. Your add-in opens a pop-up window to start the authentication flow. Add-ins cannot start authentication flows directly because add-ins, depending on the platform used, may run in an IFRAME. For security reasons, OAuth sign-in pages can't be displayed in an IFRAME.</span></span> 
    
- <span data-ttu-id="948f0-p106">Un serveur web héberge le code de votre complément. Cet exemple de code utilise un serveur de base de données exécuté sur le serveur web pour stocker le jeton d’accès de l’utilisateur. Il est nécessaire de conserver le jeton d’accès pour que, une fois l’authentification terminée dans la fenêtre contextuelle, les pages principales du complément puissent utiliser les mêmes jetons pour accéder aux données à partir du service en ligne. Il est nécessaire d’enregistrer les jetons sur le serveur car vous ne pouvez pas vous fier aux informations transmises depuis le complément ou la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="948f0-p106">A web server hosts your add-in's code. This code sample uses a database server running on the web server to store the user's access token. Persisting the access token is necessary so that after authentication completes using the pop-up window, the main add-in's pages can use the same tokens to access data from the online service. Saving the tokens by using server-side options is necessary because you can't rely on information passed from the add-in or the pop-up.</span></span>
    
- <span data-ttu-id="948f0-126">Le fournisseur OAuth 2.0 effectue l’authentification de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="948f0-126">The OAuth 2.0 provider performs user authentication.</span></span>
    

    
> [!IMPORTANT]
> <span data-ttu-id="948f0-p107">Les jetons d’accès ne peuvent pas être renvoyés au volet Office, mais ils peuvent être utilisés sur le serveur. Dans cet exemple de code, les jetons d’accès sont stockés dans la base de données pendant deux minutes. Passé ce délai, ils sont purgés de la base de données et les utilisateurs sont invités à se ré-authentifier. Avant de changer ce délai dans vos paramètres, pensez aux risques de sécurité que peut poser le stockage de jetons d’accès dans une base de données pendant plus de deux minutes.</span><span class="sxs-lookup"><span data-stu-id="948f0-p107">Access tokens can't be returned to the task pane, but they can be used on the server. In this code sample, the access tokens are stored in the database for 2 minutes. After 2 minutes, tokens are purged from the database and users are prompted to re-authenticate. Before changing this time period in your own implementation, consider the security risks associated with storing access tokens in a database for a time period that is longer than 2 minutes.</span></span>


## <a name="step-1---start-socket-and-open-a-pop-up-window"></a><span data-ttu-id="948f0-131">Étape 1 : démarrer un socket et ouvrer une fenêtre contextuelle</span><span class="sxs-lookup"><span data-stu-id="948f0-131">Step 1 - Start socket and open a pop-up window</span></span>

<span data-ttu-id="948f0-p108">Lorsque vous exécutez cet exemple de code, un complément du volet Office apparaît dans Office. Lorsque l’utilisateur choisit un fournisseur OAuth auquel se connecter, le complément crée d’abord un socket. Cet exemple utilise un socket pour offrir à l’utilisateur une expérience conviviale au sein du complément. Le complément utilise le socket pour indiquer le succès ou l’échec de l’authentification à l’utilisateur. En utilisant un socket, l’état d’authentification est facilement mis à jour sur la page principale du complément, sans aucune intervention ou interrogation de l’utilisateur. Le segment de code suivant, extrait de routes/connect.js, montre comment démarrer le socket. Celui-ci est nommé avec l’ID de session du complément, **decodedNodeCookie**. Cet exemple de code crée le socket avec [socket.io](https://socket.io/).</span><span class="sxs-lookup"><span data-stu-id="948f0-p108">When you run this code sample, a task pane add-in displays in Office. When the user chooses an OAuth provider to log into, the add-in first creates a socket. This sample uses a socket to provide a good user experience in the add-in. The add-in uses the socket to communicate the success or failure of the authentication to the user. By using a socket, the add-in's main page is easily updated with the authentication status, and doesn't require user interaction or polling. The following code segment, taken from routes/connect.js, shows how to start the socket. The socket is named using  **decodedNodeCookie**, which is the add-in's session ID. This code sample creates the socket by using [socket.io](https://socket.io/).</span></span>


```js
io.on('connection', function (socket) {
  console.log('Socket connection established');
  var jsonCookie =
    cookie.parse(socket
      .handshake
      .headers
      .cookie);
  var decodedNodeCookie =
    cookieParser
      .signedCookie(jsonCookie.nodecookie, '<Insert a random string>');
  console.log('Decoded cookie: ' + decodedNodeCookie);
  // The session ID becomes the room name for this session.
  socket.join(decodedNodeCookie);
  io.to(decodedNodeCookie).emit('init', 'Private socket session established');
});

```

<span data-ttu-id="948f0-p109">Ensuite, le complément se connecte au socket. Le code suivant est issu de /public/javascripts/client.js.</span><span class="sxs-lookup"><span data-stu-id="948f0-p109">Next, the add-in connects to the socket. The following code can be found in /public/javascripts/client.js.</span></span>




```js
var socket = io.connect('https://localhost:3001', { secure: true });
```

<span data-ttu-id="948f0-p110">Enfin, le complément ouvre une fenêtre contextuelle sur l’ordinateur de l’utilisateur en utilisant **window.open**. Lors de l’exécution de **window.open**, assurez-vous que l’URI de redirection et l’ID de session du complément sont bien transmis dans l’URL. L’ID de session du complément permet d’identifier le socket à utiliser lors de l’envoi des informations d’état d’authentification vers l’interface utilisateur du complément. Le segment de code suivant est issu de views/index.jade.</span><span class="sxs-lookup"><span data-stu-id="948f0-p110">Next, the add-in opens a pop-up window on the user's computer using  **window.open**. When running  **window.open**, ensure the redirect URI and session ID of the add-in is passed in the URL. The session ID of the add-in is used to identify the socket to use when sending authentication status information to the add-in's UI. The following code segment can be found in views/index.jade.</span></span>




```js
onclick="window.open('/connect/azure/#{sessionID}', 'AuthPopup', 'width=500,height=500,centerscreen=1,menubar=0,toolbar=0,location=0,personalbar=0,status=0,titlebar=0,dialog=1')")
```


## <a name="steps-2-amp-3---start-the-authentication-flow-and-show-the-sign-in-page"></a><span data-ttu-id="948f0-146">Étapes 2 &amp; 3 : lancer le flux d’authentification et afficher la page de connexion</span><span class="sxs-lookup"><span data-stu-id="948f0-146">Steps 2 &amp; 3 - Start the authentication flow and show the sign-in page</span></span>

<span data-ttu-id="948f0-p111">Le complément doit démarrer le flux d’authentification. Le segment de code ci-dessous utilise la bibliothèque OAuth Passport. Lors du démarrage du flux d’authentification, veillez à bien transmettre l’URL d’autorisation du fournisseur OAuth et l’ID de session du complément. L’ID de session du complément doit être transmis dans le paramètre d’état. La fenêtre contextuelle affiche maintenant la page de connexion du fournisseur OAuth et les utilisateurs peuvent se connecter.</span><span class="sxs-lookup"><span data-stu-id="948f0-p111">The add-in must start the authentication flow. The code segment below uses the Passport OAuth library. When starting the authentication flow, ensure that you pass the authorization URL of the OAuth provider, and the session ID of the add-in. The session ID of the add-in must be passed in the state parameter. The pop-up window now displays the OAuth provider's sign-in page so that users can sign in.</span></span>


```js
router.get('/azure/:sessionID', function(req, res, next) { 
   passport.authenticate( 
     'azure',  
     { state: req.params.sessionID }, 

```


## <a name="steps-4-5-amp-6---user-signs-in-and-web-server-receives-tokens"></a><span data-ttu-id="948f0-152">Étapes 4, 5 &amp; 6 : authentifier l’utilisateur et stocker les jetons sur le serveur web</span><span class="sxs-lookup"><span data-stu-id="948f0-152">Steps 4, 5 &amp; 6 - User signs in and web server receives tokens</span></span>

 <span data-ttu-id="948f0-p112">Lorsque la connexion réussit, un jeton d’accès, un jeton d’actualisation et le paramètre d’état sont renvoyés au complément. Le paramètre d’état contient l’ID de session, qui est utilisé pour envoyer des informations sur l’état d’authentification au socket, lors de l’étape 7. Le segment de code suivant, issu de app.js, stocke les jetons d’accès dans la base de données.</span><span class="sxs-lookup"><span data-stu-id="948f0-p112">After a successful sign-in, an access token, refresh token, and state parameter are returned to the add-in. The state parameter contains the session ID, which is used to send authentication status information to the socket in step 7. The following code segment, taken from app.js, stores the access token in the database.</span></span>


```js
  dbHelperInstance.insertDoc(userData, null, 
         function (err, body) { 
           if (!err) { 
             console.log("Inserted session entry [" + userData.sessid + "] id: " + body.id); 
           } 
           done(err, userData); 
         }); 

```


## <a name="step-7---show-authentication-information-in-the-add-ins-ui"></a><span data-ttu-id="948f0-156">Étape 7 : afficher les informations d’authentification dans l’interface utilisateur du complément</span><span class="sxs-lookup"><span data-stu-id="948f0-156">Step 7 - Show authentication information in the add-in's UI</span></span>

<span data-ttu-id="948f0-p113">Le segment de code suivant, issu de connect.js, met à jour l’interface utilisateur du complément avec les informations de l’état d’authentification. L’interface utilisateur du complément est mise à jour à l’aide du socket qui a été créé à l’étape 1.</span><span class="sxs-lookup"><span data-stu-id="948f0-p113">The following code segment, taken from connect.js, updates the add-in's UI with the authentication status information. The add-in's UI is updated by using the socket that was created in step 1.</span></span>


```js
  
       io.to(user.sessid).emit('auth_success', providers); 
       next(); 

```


## <a name="see-also"></a><span data-ttu-id="948f0-159">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="948f0-159">See also</span></span>

- [<span data-ttu-id="948f0-160">Exemple d’authentification du complément Office sur le serveur pour Node.js</span><span class="sxs-lookup"><span data-stu-id="948f0-160">Office Add-in Server Authentication Sample for Node.js</span></span>](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth/blob/master/README.md)
    
