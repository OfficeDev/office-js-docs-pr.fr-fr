---
title: Développer votre add-in Office pour qu’il fonctionne avec itp lors de l’utilisation de cookies tiers
description: Utilisation des modules itp et des add-ins Office lors de l’utilisation de cookies tiers
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: e66fc25e1dc0f3a93fdf38c1d0c099d3a68459d3
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178040"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a><span data-ttu-id="38fbc-103">Développer votre add-in Office pour qu’il fonctionne avec itp lors de l’utilisation de cookies tiers</span><span class="sxs-lookup"><span data-stu-id="38fbc-103">Develop your Office Add-in to work with ITP when using third-party cookies</span></span>

<span data-ttu-id="38fbc-104">Si votre add-in Office nécessite des cookies tiers, ces cookies sont bloqués si la prévention du suivi intelligent (ITP) est utilisée par le runtime du navigateur qui a chargé votre add-in.</span><span class="sxs-lookup"><span data-stu-id="38fbc-104">If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in.</span></span> <span data-ttu-id="38fbc-105">Vous pouvez utiliser des cookies tiers pour authentifier les utilisateurs ou pour d’autres scénarios, tels que le stockage des paramètres.</span><span class="sxs-lookup"><span data-stu-id="38fbc-105">You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.</span></span>

<span data-ttu-id="38fbc-106">Si votre add-in Office et votre site web doivent s’appuyer sur des cookies tiers, utilisez les étapes suivantes pour utiliser itp :</span><span class="sxs-lookup"><span data-stu-id="38fbc-106">If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:</span></span>

1. <span data-ttu-id="38fbc-107">Configurer [l’autorisation OAuth 2.0](https://tools.ietf.org/html/rfc6749)de sorte que le domaine d’authentification (dans votre cas, le tiers qui attend des cookies) a transmis un jeton d’autorisation à votre   site web.</span><span class="sxs-lookup"><span data-stu-id="38fbc-107">Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website.</span></span> <span data-ttu-id="38fbc-108">Utilisez le jeton pour établir une session de connexion tierce avec un cookie Sécurisé et [HttpOnly](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)de jeu de serveurs.</span><span class="sxs-lookup"><span data-stu-id="38fbc-108">Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).</span></span>
2. <span data-ttu-id="38fbc-109">Utilisez [l’API d’accès](https://webkit.org/blog/8124/introducing-storage-access-api/)au stockage pour que le tiers puisse demander l’autorisation d’accéder à ses   cookies tiers.</span><span class="sxs-lookup"><span data-stu-id="38fbc-109">Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies.</span></span> <span data-ttu-id="38fbc-110">Les versions actuelles d’Office sur Mac et d’Office sur le web la prise en charge de cette API.</span><span class="sxs-lookup"><span data-stu-id="38fbc-110">Current versions of Office on Mac and Office on the web both support this API.</span></span>
    > [!NOTE]
    > <span data-ttu-id="38fbc-111">Si vous utilisez des cookies à des fins autres que l’authentification, envisagez d’utiliser `localStorage` pour votre scénario.</span><span class="sxs-lookup"><span data-stu-id="38fbc-111">If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.</span></span>

<span data-ttu-id="38fbc-112">L’exemple de code suivant montre comment utiliser l’API d’accès au stockage :</span><span class="sxs-lookup"><span data-stu-id="38fbc-112">The following code sample shows how to use the Storage Access API:</span></span>

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

## <a name="about-itp-and-third-party-cookies"></a><span data-ttu-id="38fbc-113">À propos des cookies itp et tiers</span><span class="sxs-lookup"><span data-stu-id="38fbc-113">About ITP and third-party cookies</span></span>

<span data-ttu-id="38fbc-114">Les cookies tiers sont des cookies chargés dans un iframe, où le domaine est différent de l’image de niveau supérieur.</span><span class="sxs-lookup"><span data-stu-id="38fbc-114">Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame.</span></span> <span data-ttu-id="38fbc-115">Le programme itp peut affecter des scénarios d’authentification complexes, où une boîte de dialogue popup est utilisée pour entrer les informations d’identification, puis l’accès au cookie est nécessaire à un iframe de compl?ment pour terminer le flux d’authentification.</span><span class="sxs-lookup"><span data-stu-id="38fbc-115">ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow.</span></span> <span data-ttu-id="38fbc-116">Le service ITP peut également affecter les scénarios d’authentification sans fil, où vous avez déjà utilisé une boîte de dialogue popup pour s’authentifier, mais l’utilisation ultérieure du module de authentification tente de s’authentifier par le biais d’un iframe masqué.</span><span class="sxs-lookup"><span data-stu-id="38fbc-116">ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.</span></span>

<span data-ttu-id="38fbc-117">Lors du développement de add-ins Office sur Mac, l’accès aux cookies tiers est bloqué par le SDK MacOS Big Sur.</span><span class="sxs-lookup"><span data-stu-id="38fbc-117">When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK.</span></span> <span data-ttu-id="38fbc-118">Cela est dû au fait que webKit ITP est activé par défaut sur le navigateur Safari et que WKWebview bloque tous les cookies tiers.</span><span class="sxs-lookup"><span data-stu-id="38fbc-118">This is because WebKit ITP is enabled by default on the Safari browser, and WKWebview blocks all third-party cookies.</span></span> <span data-ttu-id="38fbc-119">Office sur Mac version 16.44 ou ultérieure est intégré au SDK MacOS Big Sur.</span><span class="sxs-lookup"><span data-stu-id="38fbc-119">Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.</span></span>

<span data-ttu-id="38fbc-120">Dans le navigateur Safari, les utilisateurs finaux peuvent activer la case à cocher Empêcher le suivi entre sites sous Confidentialité des préférences pour désactiver   >   l’itp.</span><span class="sxs-lookup"><span data-stu-id="38fbc-120">In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP.</span></span> <span data-ttu-id="38fbc-121">Toutefois, itp ne peut pas être désactivé pour le contrôle WebKit2 incorporé.</span><span class="sxs-lookup"><span data-stu-id="38fbc-121">However, ITP cannot be turned off for the embedded WebKit2 control.</span></span>

## <a name="see-also"></a><span data-ttu-id="38fbc-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="38fbc-122">See also</span></span>

- [<span data-ttu-id="38fbc-123">Gérer l’itp dans Safari et d’autres navigateurs où les cookies tiers sont bloqués</span><span class="sxs-lookup"><span data-stu-id="38fbc-123">Handle ITP in Safari and other browsers where third-party cookies are blocked</span></span>](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [<span data-ttu-id="38fbc-124">Prévention du suivi dans WebKit</span><span class="sxs-lookup"><span data-stu-id="38fbc-124">Tracking Prevention in WebKit</span></span>](https://webkit.org/tracking-prevention/)
- [<span data-ttu-id="38fbc-125">Chrome « Bac à sable (sandbox) de confidentialité »</span><span class="sxs-lookup"><span data-stu-id="38fbc-125">Chrome’s “Privacy Sandbox”</span></span>](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [<span data-ttu-id="38fbc-126">Présentation de l’API d’accès au stockage</span><span class="sxs-lookup"><span data-stu-id="38fbc-126">Introducing the Storage Access API</span></span>](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)