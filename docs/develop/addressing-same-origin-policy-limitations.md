---
title: Résolutions des limites de stratégie d’origine identique dans les compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5aa329eb3f073f3544d8446683debed3239fd00
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270599"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Résolutions des limites de stratégie d’origine identique dans les compléments Office


La stratégie de même origine appliquée par le navigateur empêche un script chargé à partir d’un domaine d’obtenir ou de manipuler les propriétés d’une page web issue d’un autre domaine. Cela signifie que, par défaut, le domaine d’une URL demandée doit correspondre au domaine de la page web actuelle. Par exemple, cette stratégie empêche une page web d’un domaine d’effectuer des appels de service web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) à un domaine autre que celui où elle est hébergée.

Comme les Compléments Office sont hébergés dans un contrôle de navigateur, la stratégie de même origine s’applique également aux scripts exécutés dans leurs pages web.

Il existe de nombreuses manières d’annuler le complément de la stratégie de même origine lorsque vous développez des compléments :

- Utilisation de JSON/P pour un accès anonyme. 
    
- Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons.
    
- Utilisation du partage de ressources cross-origin (CORS).
    
- Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE.
    

## <a name="using-jsonp-for-anonymous-access"></a>Utilisation de JSON/P pour un accès anonyme


Une façon de contourner cette limitation consiste à utiliser JSON/P afin de fournir un proxy pour le service web. Pour ce faire, incluez une balise `script` avec un attribut `src` qui pointe vers un script hébergé sur n’importe quel domaine. Vous pouvez créer les balises `script` par programmation, créer dynamiquement l’URL vers laquelle pointer l’attribut `src`, puis passer des paramètres à l’URL au moyen de paramètres de requête URI. Les fournisseurs de services web créent et hébergent du code JavaScript sur des URL spécifiques et renvoient des scripts différents selon les paramètres de requête URI. Ces scripts s’exécutent ensuite là où ils sont insérés et fonctionnent comme prévu.

L’exemple suivant illustre JSON/P utilisant une technique qui fonctionne dans n’importe quel Complément Office.

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a>Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons


Une autre manière de résoudre les limitations de la stratégie de même origine consiste à implémenter la page web du complément sous la forme d’une page ASP qui utilise OAuth ou met en cache les informations d’identification dans des cookies.

Pour un exemple de code côté serveur qui illustre comment utiliser l’objet `Cookie` dans `System.Net` pour obtenir et définir des valeurs de cookie, voir la propriété [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2).


## <a name="using-cross-origin-resource-sharing-cors"></a>Utilisation du partage de ressources cross-origin (CORS)


Pour un exemple de la fonctionnalité de partage de ressources cross-origin de [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), voir la section « Partage de ressources cross-origin (CORS) » de [Nouvelles astuces dans XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE


Pour un exemple de construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE, reportez-vous à [Messagerie entre fenêtres](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
    
