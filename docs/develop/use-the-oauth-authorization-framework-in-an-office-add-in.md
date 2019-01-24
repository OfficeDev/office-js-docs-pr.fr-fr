---
title: Utilisation d’OAuth dans un complément Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 9689c92bb5b118f9d45c1805a094b4243b4e6049
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389276"
---
# <a name="use-the-oauth-authorization-framework-in-an-office-add-in"></a>Utilisation d’OAuth dans un complément Office

OAuth est le standard d’authentification ouvert utilisé par les fournisseurs de services en ligne (Office 365, Facebook, Google, SalesForce, LinkedIn, etc.) pour procéder à l’authentification des utilisateurs. OAuth est le protocole d’authentification par défaut pour Azure et Office 365. OAuth est utilisé par les entreprises et les consommateurs.

Les fournisseurs de services en ligne peuvent proposer des API publiques qui se connectent avec REST. Les développeurs peuvent utiliser ces API publiques dans leurs compléments Office pour lire ou écrire des données au fournisseur de services en ligne. Intégrer les données des fournisseurs de services en ligne dans un complément augmente sa valeur et entraîne de meilleurs taux d’adoption. Lorsque vous utilisez ces API dans votre complément, les utilisateurs doivent s’authentifier avec OAuth.

Cette rubrique explique comment mettre en œuvre un flux d’authentification dans votre complément pour procéder à l’authentification des utilisateurs. Les segments de code inclus dans cette rubrique sont extraits de l’exemple de code [Office-Add-in-NodeJS-ServerAuth](https://github.com/OfficeDev/Office-Add-in-NodeJS-ServerAuth).

> [!NOTE]
> Pour des raisons de sécurité, les navigateurs ne sont pas autorisés à afficher les pages de connexion dans un IFrame. Selon la version d’Office utilisée par vos clients, notamment les versions web, votre complément s’affiche dans un IFrame. C’est pourquoi il est indispensable de savoir comment gérer le flux d’authentification.  

Le schéma suivant présente les composants nécessaires, ainsi que le flux d’événements qui se produisent lors de l’implémentation de l’authentification dans votre complément.

![Réalisation d’une authentification OAuth dans un complément Office](../images/oauth-in-office-add-in.png)

Le schéma illustre la façon dont les composants requis suivants sont utilisés :


- Office exécute un complément de volet Office sur l’ordinateur de l’utilisateur. Votre complément ouvre une fenêtre contextuelle pour démarrer le flux d’authentification. Les compléments ne peuvent pas démarrer les flux d’authentification directement, car ils sont exécutés dans un IFRAME (selon la plateforme utilisée). Pour des raisons de sécurité, les pages de connexion OAuth ne peuvent pas être affichées dans un IFRAME. 
    
- Un serveur web héberge le code de votre complément. Cet exemple de code utilise un serveur de base de données exécuté sur le serveur web pour stocker le jeton d’accès de l’utilisateur. Il est nécessaire de conserver le jeton d’accès pour que, une fois l’authentification terminée dans la fenêtre contextuelle, les pages principales du complément puissent utiliser les mêmes jetons pour accéder aux données à partir du service en ligne. Il est nécessaire d’enregistrer les jetons sur le serveur car vous ne pouvez pas vous fier aux informations transmises depuis le complément ou la fenêtre contextuelle.
    
- Le fournisseur OAuth 2.0 effectue l’authentification de l’utilisateur.
    

    
> [!IMPORTANT]
> Les jetons d’accès ne peuvent pas être renvoyés au volet Office, mais ils peuvent être utilisés sur le serveur. Dans cet exemple de code, les jetons d’accès sont stockés dans la base de données pendant deux minutes. Passé ce délai, ils sont purgés de la base de données et les utilisateurs sont invités à se ré-authentifier. Avant de changer ce délai dans vos paramètres, pensez aux risques de sécurité que peut poser le stockage de jetons d’accès dans une base de données pendant plus de deux minutes.


## <a name="step-1---start-socket-and-open-a-pop-up-window"></a>Étape 1 : démarrer un socket et ouvrer une fenêtre contextuelle

Lorsque vous exécutez cet exemple de code, un complément du volet Office apparaît dans Office. Lorsque l’utilisateur choisit un fournisseur OAuth auquel se connecter, le complément crée d’abord un socket. Cet exemple utilise un socket pour offrir à l’utilisateur une expérience conviviale au sein du complément. Le complément utilise le socket pour indiquer le succès ou l’échec de l’authentification à l’utilisateur. En utilisant un socket, l’état d’authentification est facilement mis à jour sur la page principale du complément, sans aucune intervention ou interrogation de l’utilisateur. Le segment de code suivant, extrait de routes/connect.js, montre comment démarrer le socket. Celui-ci est nommé avec l’ID de session du complément, **decodedNodeCookie**. Cet exemple de code crée le socket avec [socket.io](https://socket.io/).


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

Ensuite, le complément se connecte au socket. Le code suivant est issu de /public/javascripts/client.js.




```js
var socket = io.connect('https://localhost:3001', { secure: true });
```

Enfin, le complément ouvre une fenêtre contextuelle sur l’ordinateur de l’utilisateur en utilisant **window.open**. Lors de l’exécution de **window.open**, assurez-vous que l’URI de redirection et l’ID de session du complément sont bien transmis dans l’URL. L’ID de session du complément permet d’identifier le socket à utiliser lors de l’envoi des informations d’état d’authentification vers l’interface utilisateur du complément. Le segment de code suivant est issu de views/index.jade.




```js
onclick="window.open('/connect/azure/#{sessionID}', 'AuthPopup', 'width=500,height=500,centerscreen=1,menubar=0,toolbar=0,location=0,personalbar=0,status=0,titlebar=0,dialog=1')")
```


## <a name="steps-2-amp-3---start-the-authentication-flow-and-show-the-sign-in-page"></a>Étapes 2 &amp; 3 : lancer le flux d’authentification et afficher la page de connexion

Le complément doit démarrer le flux d’authentification. Le segment de code ci-dessous utilise la bibliothèque OAuth Passport. Lors du démarrage du flux d’authentification, veillez à bien transmettre l’URL d’autorisation du fournisseur OAuth et l’ID de session du complément. L’ID de session du complément doit être transmis dans le paramètre d’état. La fenêtre contextuelle affiche maintenant la page de connexion du fournisseur OAuth et les utilisateurs peuvent se connecter.


```js
router.get('/azure/:sessionID', function(req, res, next) { 
   passport.authenticate( 
     'azure',  
     { state: req.params.sessionID }, 

```


## <a name="steps-4-5-amp-6---user-signs-in-and-web-server-receives-tokens"></a>Étapes 4, 5 &amp; 6 : authentifier l’utilisateur et stocker les jetons sur le serveur web

 Lorsque la connexion réussit, un jeton d’accès, un jeton d’actualisation et le paramètre d’état sont renvoyés au complément. Le paramètre d’état contient l’ID de session, qui est utilisé pour envoyer des informations sur l’état d’authentification au socket, lors de l’étape 7. Le segment de code suivant, issu de app.js, stocke les jetons d’accès dans la base de données.


```js
  dbHelperInstance.insertDoc(userData, null, 
         function (err, body) { 
           if (!err) { 
             console.log("Inserted session entry [" + userData.sessid + "] id: " + body.id); 
           } 
           done(err, userData); 
         }); 

```


## <a name="step-7---show-authentication-information-in-the-add-ins-ui"></a>Étape 7 : afficher les informations d’authentification dans l’interface utilisateur du complément

Le segment de code suivant, issu de connect.js, met à jour l’interface utilisateur du complément avec les informations de l’état d’authentification. L’interface utilisateur du complément est mise à jour à l’aide du socket qui a été créé à l’étape 1.


```js
  
       io.to(user.sessid).emit('auth_success', providers); 
       next(); 

```


## <a name="see-also"></a>Voir aussi

- [Exemple d’authentification du complément Office sur le serveur pour Node.js](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth/blob/master/README.md)
    
