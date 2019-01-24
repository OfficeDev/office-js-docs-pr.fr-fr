---
title: Résolutions des limites de stratégie d’origine identique dans les compléments Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 75bc42cd7d2a7acc8cb57ee08807a8486e21f467
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387754"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="52b51-102">Résolutions des limites de stratégie d’origine identique dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="52b51-102">Addressing same-origin policy limitations in Office Add-ins</span></span>


<span data-ttu-id="52b51-p101">La stratégie de même origine appliquée par le navigateur empêche un script chargé à partir d’un domaine d’obtenir ou de manipuler les propriétés d’une page web issue d’un autre domaine. Cela signifie que, par défaut, le domaine d’une URL demandée doit correspondre au domaine de la page web actuelle. Par exemple, cette stratégie empêche une page web d’un domaine d’effectuer des appels de service web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) à un domaine autre que celui où elle est hébergée.</span><span class="sxs-lookup"><span data-stu-id="52b51-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="52b51-106">Comme les Compléments Office sont hébergés dans un contrôle de navigateur, la stratégie de même origine s’applique également aux scripts exécutés dans leurs pages web.</span><span class="sxs-lookup"><span data-stu-id="52b51-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="52b51-107">Il existe de nombreuses manières d’annuler le complément de la stratégie de même origine lorsque vous développez des compléments :</span><span class="sxs-lookup"><span data-stu-id="52b51-107">To overcome same-origin policy enforcement when you develop add-ins, you can:</span></span>

- <span data-ttu-id="52b51-108">Utilisation de JSON/P pour un accès anonyme.</span><span class="sxs-lookup"><span data-stu-id="52b51-108">Use JSON/P for anonymous access.</span></span> 
    
- <span data-ttu-id="52b51-109">Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons.</span><span class="sxs-lookup"><span data-stu-id="52b51-109">Implement server-side script using a token-based authentication scheme.</span></span>
    
- <span data-ttu-id="52b51-110">Utilisation du partage de ressources cross-origin (CORS).</span><span class="sxs-lookup"><span data-stu-id="52b51-110">Using cross-origin resource sharing (CORS).</span></span>
    
- <span data-ttu-id="52b51-111">Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE.</span><span class="sxs-lookup"><span data-stu-id="52b51-111">Build your own proxy using IFRAME and POST MESSAGE.</span></span>
    

## <a name="using-jsonp-for-anonymous-access"></a><span data-ttu-id="52b51-112">Utilisation de JSON/P pour un accès anonyme</span><span class="sxs-lookup"><span data-stu-id="52b51-112">Using JSON/P for anonymous access</span></span>


<span data-ttu-id="52b51-p102">Une façon de contourner cette limitation consiste à utiliser JSON/P afin de fournir un proxy pour le service web. Pour ce faire, incluez une balise `script` avec un attribut `src` qui pointe vers un script hébergé sur n’importe quel domaine. Vous pouvez créer les balises `script` par programmation, créer dynamiquement l’URL vers laquelle pointer l’attribut `src`, puis passer des paramètres à l’URL au moyen de paramètres de requête URI. Les fournisseurs de services web créent et hébergent du code JavaScript sur des URL spécifiques et renvoient des scripts différents selon les paramètres de requête URI. Ces scripts s’exécutent ensuite là où ils sont insérés et fonctionnent comme prévu.</span><span class="sxs-lookup"><span data-stu-id="52b51-p102">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="52b51-118">L’exemple suivant illustre JSON/P utilisant une technique qui fonctionne dans n’importe quel Complément Office.</span><span class="sxs-lookup"><span data-stu-id="52b51-118">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a><span data-ttu-id="52b51-119">Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons</span><span class="sxs-lookup"><span data-stu-id="52b51-119">Implementing server-side script using a token-based authentication scheme</span></span>


<span data-ttu-id="52b51-120">Une autre manière de résoudre les limitations de la stratégie de même origine consiste à implémenter la page web du complément sous la forme d’une page ASP qui utilise OAuth ou met en cache les informations d’identification dans des cookies.</span><span class="sxs-lookup"><span data-stu-id="52b51-120">Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.</span></span>

<span data-ttu-id="52b51-121">Pour un exemple de code côté serveur qui illustre comment utiliser l’objet `Cookie` dans `System.Net` pour obtenir et définir des valeurs de cookie, voir la propriété [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2).</span><span class="sxs-lookup"><span data-stu-id="52b51-121">For an example of server-side code that shows how to use the  `Cookie` object in `System.Net` to get and set cookie values, see the [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) property.</span></span>


## <a name="using-cross-origin-resource-sharing-cors"></a><span data-ttu-id="52b51-122">Utilisation du partage de ressources cross-origin (CORS)</span><span class="sxs-lookup"><span data-stu-id="52b51-122">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="52b51-123">Pour un exemple de la fonctionnalité de partage de ressources cross-origin de [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), voir la section « Partage de ressources cross-origin (CORS) » de [Nouvelles astuces dans XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="52b51-123">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a><span data-ttu-id="52b51-124">Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE</span><span class="sxs-lookup"><span data-stu-id="52b51-124">Building your own proxy using IFRAME and POST MESSAGE</span></span>


<span data-ttu-id="52b51-125">Pour un exemple de construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE, reportez-vous à [Messagerie entre fenêtres](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="52b51-125">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="52b51-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="52b51-126">See also</span></span>

- [<span data-ttu-id="52b51-127">Confidentialité et sécurité pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="52b51-127">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
