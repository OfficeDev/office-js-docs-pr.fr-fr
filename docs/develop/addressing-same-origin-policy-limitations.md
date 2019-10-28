---
title: Résolutions des limites de stratégie d’origine identique dans les compléments Office
description: ''
ms.date: 10/17/2019
localization_priority: Priority
ms.openlocfilehash: 353dfc40df79f74614bb2eab026a65c54c3f16b9
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37681836"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="05d6d-102">Résolutions des limites de stratégie d’origine identique dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="05d6d-102">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="05d6d-p101">La stratégie de même origine appliquée par le navigateur empêche un script chargé à partir d’un domaine d’obtenir ou de manipuler les propriétés d’une page web issue d’un autre domaine. Cela signifie que, par défaut, le domaine d’une URL demandée doit correspondre au domaine de la page web actuelle. Par exemple, cette stratégie empêche une page web d’un domaine d’effectuer des appels de service web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) à un domaine autre que celui où elle est hébergée.</span><span class="sxs-lookup"><span data-stu-id="05d6d-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="05d6d-106">Comme les Compléments Office sont hébergés dans un contrôle de navigateur, la stratégie de même origine s’applique également aux scripts exécutés dans leurs pages web.</span><span class="sxs-lookup"><span data-stu-id="05d6d-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="05d6d-107">La stratégie de la même origine peut être un handicap inutile dans de nombreuses situations, par exemple, quand une application web héberge du contenu et des API au sein de plusieurs sous-domaines.</span><span class="sxs-lookup"><span data-stu-id="05d6d-107">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="05d6d-108">Il existe quelques techniques permettant de surmonter le renforcement de la stratégie de la même origine.</span><span class="sxs-lookup"><span data-stu-id="05d6d-108">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="05d6d-109">Cet article peut fournir uniquement l’introduction la plus courte à certains d'entre eux.</span><span class="sxs-lookup"><span data-stu-id="05d6d-109">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="05d6d-110">Utilisez des liens fournis pour commencer à utiliser vos recherches des techniques suivantes.</span><span class="sxs-lookup"><span data-stu-id="05d6d-110">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="05d6d-111">Utilisation de JSONP pour un accès anonyme</span><span class="sxs-lookup"><span data-stu-id="05d6d-111">Use JSON/P for anonymous access</span></span>

<span data-ttu-id="05d6d-112">Une façon de contourner cette limitation consiste à utiliser [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) afin de fournir un proxy pour le service web.</span><span class="sxs-lookup"><span data-stu-id="05d6d-112">One way to overcome same-origin policy limitations is to use [JSON/P](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service.</span></span> <span data-ttu-id="05d6d-113">Pour ce faire, incluez une balise `script` avec un attribut `src` qui pointe vers un script hébergé sur n’importe quel domaine.</span><span class="sxs-lookup"><span data-stu-id="05d6d-113">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="05d6d-114">Vous pouvez créer les balises `script` par programmation, créer dynamiquement l’URL vers laquelle pointer l’attribut `src`, puis passer des paramètres à l’URL au moyen de paramètres de requête URI.</span><span class="sxs-lookup"><span data-stu-id="05d6d-114">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="05d6d-115">Les fournisseurs de services web créent et hébergent du code JavaScript sur des URL spécifiques et renvoient des scripts différents selon les paramètres de requête URI.</span><span class="sxs-lookup"><span data-stu-id="05d6d-115">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="05d6d-116">Ces scripts s’exécutent ensuite là où ils sont insérés et fonctionnent comme prévu.</span><span class="sxs-lookup"><span data-stu-id="05d6d-116">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="05d6d-117">L’exemple suivant illustre JSONP utilisant une technique qui fonctionne dans n’importe quel Complément Office.</span><span class="sxs-lookup"><span data-stu-id="05d6d-117">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="05d6d-118">Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons.</span><span class="sxs-lookup"><span data-stu-id="05d6d-118">Implement server-side code using a token-based authorization scheme</span></span>

<span data-ttu-id="05d6d-119">Une autre méthode d’aborder les limitations spécifiques de stratégie de la même origine fournit le code côté serveur qui utilise les flux[OAuth 2.0](https://oauth.net/2/)pour activer un domaine autorisé afin d’accéder aux ressources hébergées sur un autre.</span><span class="sxs-lookup"><span data-stu-id="05d6d-119">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="05d6d-120">Utilisation du partage de ressources cross-origin (CORS)</span><span class="sxs-lookup"><span data-stu-id="05d6d-120">Use cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="05d6d-121">Pour un exemple de la fonctionnalité de partage de ressources cross-origin de [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), voir la section « Partage de ressources cross-origin (CORS) » de [Nouvelles astuces dans XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="05d6d-121">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="05d6d-122">Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE (Messagerie entre-fenêtre).</span><span class="sxs-lookup"><span data-stu-id="05d6d-122">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="05d6d-123">Pour un exemple de construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE, reportez-vous à [Messagerie entre fenêtres](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="05d6d-123">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="05d6d-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="05d6d-124">See also</span></span>

- [<span data-ttu-id="05d6d-125">Confidentialité et sécurité pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="05d6d-125">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
