---
title: Résolutions des limites de stratégie d’origine identique dans les compléments Office
description: ''
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: 2a47339bd5cc0b0bf919152b7078d5373382124f
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950445"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Résolutions des limites de stratégie d’origine identique dans les compléments Office

La stratégie de même origine appliquée par le navigateur empêche un script chargé à partir d’un domaine d’obtenir ou de manipuler les propriétés d’une page web issue d’un autre domaine. Cela signifie que, par défaut, le domaine d’une URL demandée doit correspondre au domaine de la page web actuelle. Par exemple, cette stratégie empêche une page web d’un domaine d’effectuer des appels de service web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) à un domaine autre que celui où elle est hébergée.

Comme les Compléments Office sont hébergés dans un contrôle de navigateur, la stratégie de même origine s’applique également aux scripts exécutés dans leurs pages web.

La stratégie de la même origine peut être un handicap inutile dans de nombreuses situations, par exemple, quand une application web héberge du contenu et des API au sein de plusieurs sous-domaines. Il existe quelques techniques permettant de surmonter le renforcement de la stratégie de la même origine. Cet article peut fournir uniquement l’introduction la plus courte à certains d'entre eux. Utilisez des liens fournis pour commencer à utiliser vos recherches des techniques suivantes.

## <a name="use-jsonp-for-anonymous-access"></a>Utilisation de JSONP pour un accès anonyme

Une façon de contourner cette limitation consiste à utiliser [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) afin de fournir un proxy pour le service web. Pour ce faire, incluez une balise `script` avec un attribut `src` qui pointe vers un script hébergé sur n’importe quel domaine. Vous pouvez créer les balises `script` par programmation, créer dynamiquement l’URL vers laquelle pointer l’attribut `src`, puis passer des paramètres à l’URL au moyen de paramètres de requête URI. Les fournisseurs de services web créent et hébergent du code JavaScript sur des URL spécifiques et renvoient des scripts différents selon les paramètres de requête URI. Ces scripts s’exécutent ensuite là où ils sont insérés et fonctionnent comme prévu.

L’exemple suivant illustre JSONP utilisant une technique qui fonctionne dans n’importe quel Complément Office.

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a>Implémentation d’un script coté serveur à l’aide d’un schéma d’authentification basé sur les jetons.

Une autre méthode d’aborder les limitations spécifiques de stratégie de la même origine fournit le code côté serveur qui utilise les flux[OAuth 2.0](https://oauth.net/2/)pour activer un domaine autorisé afin d’accéder aux ressources hébergées sur un autre. 


## <a name="use-cross-origin-resource-sharing-cors"></a>Utilisation du partage de ressources cross-origin (CORS)


Pour un exemple de la fonctionnalité de partage de ressources cross-origin de [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), voir la section « Partage de ressources cross-origin (CORS) » de [Nouvelles astuces dans XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a>Construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE (Messagerie entre-fenêtre).


Pour un exemple de construction de votre propre proxy à l’aide d’IFRAME et de POSTMESSAGE, reportez-vous à [Messagerie entre fenêtres](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
    
