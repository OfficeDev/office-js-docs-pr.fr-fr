---
title: Développer votre complément Office pour utiliser itp lors de l’utilisation de cookies tiers
description: Utilisation des compléments ITP et Office lors de l’utilisation de cookies tiers
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: b01051fa39441fddb2453b0bd95a0629ebf3ef65
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423089"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Développer votre complément Office pour utiliser itp lors de l’utilisation de cookies tiers

Si votre complément Office nécessite des cookies tiers, ces cookies sont bloqués si le [runtime](../testing/runtimes.md) qui a chargé votre complément utilise la prévention du suivi intelligent (ITP). Vous pouvez utiliser des cookies tiers pour authentifier les utilisateurs, ou pour d’autres scénarios, tels que le stockage des paramètres.

Si votre complément Office et votre site web doivent s’appuyer sur des cookies tiers, procédez comme suit pour utiliser itp.

1. Configurez l’autorisation  [OAuth 2.0](https://tools.ietf.org/html/rfc6749)afin que le domaine d’authentification (dans votre cas, le tiers qui attend des cookies) transfère un jeton d’autorisation à votre site web. Utilisez le jeton pour établir une session de connexion interne avec un [cookie Sécurisé et HttpOnly](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies) défini par le serveur.
1. Utilisez l’API  [d’accès au stockage](https://webkit.org/blog/8124/introducing-storage-access-api/)afin que le tiers puisse demander l’autorisation d’accéder à ses cookies internes. Les versions actuelles d’Office sur Mac et Office sur le Web prennent toutes les deux en charge cette API.
    > [!NOTE]
    > Si vous utilisez des cookies à des fins autres que l’authentification, envisagez d’utiliser `localStorage` pour votre scénario.

L’exemple de code suivant montre comment utiliser l’API d’accès au stockage.

```javascript
function displayLoginButton() {
  const button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## <a name="about-itp-and-third-party-cookies"></a>À propos des cookies ITP et tiers

Les cookies tiers sont des cookies qui sont chargés dans un iframe, où le domaine est différent du cadre de niveau supérieur. ItP peut affecter les scénarios d’authentification complexes, où une boîte de dialogue contextuelle est utilisée pour entrer des informations d’identification, puis l’accès aux cookies est nécessaire par un iframe de complément pour terminer le flux d’authentification. ItP peut également affecter les scénarios d’authentification silencieuse, où vous avez précédemment utilisé une boîte de dialogue contextuelle pour l’authentification, mais l’utilisation ultérieure du complément tente de s’authentifier via un iframe masqué.

Lors du développement de compléments Office sur Mac, l’accès aux cookies tiers est bloqué par le SDK MacOS Big Sur. Cela est dû au fait que WKWebView ITP est activé par défaut sur le navigateur Safari et que WKWebView bloque tous les cookies tiers. Office sur Mac version 16.44 ou ultérieure est intégré au SDK MacOS Big Sur.

Dans le navigateur Safari, les utilisateurs finaux peuvent activer la case à cocher **Empêcher le suivi intersites** sous **Confidentialité** des **préférences** >  pour désactiver itp. Toutefois, itp ne peut pas être désactivé pour le contrôle WKWebView incorporé.

## <a name="see-also"></a>Voir aussi

- [Gérer itp dans Safari et d’autres navigateurs où les cookies tiers sont bloqués](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Prévention du suivi dans WebKit](https://webkit.org/tracking-prevention/)
- [« Bac à sable de confidentialité » de Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Présentation de l’API d’accès au stockage](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
