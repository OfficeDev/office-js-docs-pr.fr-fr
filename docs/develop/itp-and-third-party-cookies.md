---
title: Développer votre add-in Office pour qu'il fonctionne avec le service ITP lors de l'utilisation de cookies tiers
description: Utilisation des modules itp et des add-ins Office lors de l'utilisation de cookies tiers
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 468147e923bb27638e45879104db75b99d014986
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917092"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Développer votre add-in Office pour qu'il fonctionne avec le service ITP lors de l'utilisation de cookies tiers

Si votre add-in Office nécessite des cookies tiers, ces cookies sont bloqués si la prévention du suivi intelligent (ITP) est utilisée par le runtime du navigateur qui a chargé votre add-in. Vous pouvez utiliser des cookies tiers pour authentifier les utilisateurs ou pour d'autres scénarios, tels que le stockage des paramètres.

Si votre add-in Office et votre site web doivent s'appuyer sur des cookies tiers, utilisez les étapes suivantes pour utiliser itp :

1. Configurer [l'autorisation OAuth 2.0](https://tools.ietf.org/html/rfc6749)de sorte que le domaine d'authentification (dans votre cas, le tiers qui attend des cookies) a transmis un jeton d'autorisation à votre site   web. Utilisez le jeton pour établir une session de connexion tierce avec un cookie Sécurisé et [HttpOnly](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)de jeu de serveurs.
2. Utilisez [l'API d'accès](https://webkit.org/blog/8124/introducing-storage-access-api/)au stockage pour que le tiers puisse demander l'autorisation d'accéder à ses   cookies tiers. Les versions actuelles d'Office sur Mac et d'Office sur le web la prise en charge de cette API.
    > [!NOTE]
    > Si vous utilisez des cookies à des fins autres que l'authentification, envisagez d'utiliser `localStorage` pour votre scénario.

L'exemple de code suivant montre comment utiliser l'API d'accès au stockage :

```javascript
function displayLoginButton() {
  var button = createLoginButton();
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

## <a name="about-itp-and-third-party-cookies"></a>À propos des cookies itp et tiers

Les cookies tiers sont des cookies chargés dans un iframe, où le domaine est différent de l'image de niveau supérieur. Le programme itp peut affecter des scénarios d'authentification complexes, où une boîte de dialogue popup est utilisée pour entrer les informations d'identification, puis l'accès au cookie est nécessaire à un iframe de compl?ment pour terminer le flux d'authentification. Le service ITP peut également affecter les scénarios d'authentification sans fil, où vous avez déjà utilisé une boîte de dialogue popup pour s'authentifier, mais l'utilisation ultérieure du module de authentification tente de s'authentifier par le biais d'un iframe masqué.

Lors du développement de add-ins Office sur Mac, l'accès aux cookies tiers est bloqué par le SDK MacOS Big Sur. En effet, WKWebView ITP est activé par défaut sur le navigateur Safari et WKWebView bloque tous les cookies tiers. Office sur Mac version 16.44 ou ultérieure est intégré au SDK MacOS Big Sur.

Dans le navigateur Safari, les utilisateurs finaux peuvent activer la case à cocher Empêcher le suivi entre sites sous Confidentialité des préférences pour désactiver   >   l'itp. Toutefois, itp ne peut pas être désactivé pour le contrôle WKWebView incorporé.

## <a name="see-also"></a>Voir aussi

- [Gérer l'itp dans Safari et d'autres navigateurs où les cookies tiers sont bloqués](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Prévention du suivi dans WebKit](https://webkit.org/tracking-prevention/)
- [Chrome « Bac à sable (sandbox) de confidentialité »](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Présentation de l'API d'accès au stockage](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)