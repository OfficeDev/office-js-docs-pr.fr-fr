---
title: Authentifier un utilisateur avec un jeton identité dans un complément
description: Découvrez comment utiliser le jeton d’identité fourni par un complément Outlook pour implémenter l’authentification unique SSO dans votre service.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4134aa8ff21262f2f384d141db002b56a4a32f0a
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166143"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a><span data-ttu-id="9f021-103">Authentifier un utilisateur avec un jeton d’identité pour Exchange</span><span class="sxs-lookup"><span data-stu-id="9f021-103">Authenticate a user with an identity token for Exchange</span></span>

<span data-ttu-id="9f021-104">Les jetons d’identité d’utilisateur Exchange sont un moyen pour vos compléments d'identifier leurs utilisateurs de manière unique.
</span><span class="sxs-lookup"><span data-stu-id="9f021-104">Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user.</span></span> <span data-ttu-id="9f021-105">En établissant l’identité des utilisateurs, vous pouvez implémenter pour votre service principal un modèle d’authentification unique (SSO) qui permet aux utilisateurs de compléments Outlook de s’y connecter sans ouvrir de session.
</span><span class="sxs-lookup"><span data-stu-id="9f021-105">By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without logging in.</span></span> <span data-ttu-id="9f021-106">Pour plus d’informations sur l’utilisation de ce type de jeton, voir [Jeton d’identité d’utilisateur Exchange](authentication.md#exchange-user-identity-token).</span><span class="sxs-lookup"><span data-stu-id="9f021-106">See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type.</span></span> <span data-ttu-id="9f021-107">Dans cet article, nous allons examiner une méthode simple pour authentifier un utilisateur sur votre back end à l’aide d’un jeton d’identité Exchange.
</span><span class="sxs-lookup"><span data-stu-id="9f021-107">In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9f021-108">Il s’agit tout simplement d’un exemple d’implémentation d’une authentification unique.</span><span class="sxs-lookup"><span data-stu-id="9f021-108">This is just a simple example of an SSO implementation.</span></span> <span data-ttu-id="9f021-109">Comme toujours, lorsqu’il est question d’identité et d’authentification, vous devez vous assurer que votre code respecte les exigences en matière de sécurité de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="9f021-109">As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.</span></span>

## <a name="send-the-id-token-with-each-request"></a><span data-ttu-id="9f021-110">Envoyer le jeton d’ID avec chaque requête</span><span class="sxs-lookup"><span data-stu-id="9f021-110">Send the ID token with each request</span></span>

<span data-ttu-id="9f021-111">La première étape concerne votre complément qui doit obtenir du serveur le jeton d’identité d’utilisateur Exchange en appelant la méthode [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="9f021-111">The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span> <span data-ttu-id="9f021-112">Le complément envoie ensuite ce jeton avec chaque requête effectuée à votre serveur principal.</span><span class="sxs-lookup"><span data-stu-id="9f021-112">Then the add-in sends this token with every request it makes to your back-end.</span></span> <span data-ttu-id="9f021-113">Cela peut se faire dans un en-tête ou dans le corps de la requête.</span><span class="sxs-lookup"><span data-stu-id="9f021-113">This could be in a header, or as part of the request body.</span></span>

## <a name="validate-the-token"></a><span data-ttu-id="9f021-114">Valider le jeton</span><span class="sxs-lookup"><span data-stu-id="9f021-114">Validate the token</span></span>

<span data-ttu-id="9f021-115">Le serveur principal DOIT valider le jeton avant de l’accepter.</span><span class="sxs-lookup"><span data-stu-id="9f021-115">The back-end MUST validate the token before accepting it.</span></span> <span data-ttu-id="9f021-116">Il s’agit d’une étape importante pour garantir que le jeton a été émis par le serveur Exchange de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9f021-116">This is an important step to ensure that the token was issued by the user's Exchange server.</span></span> <span data-ttu-id="9f021-117">Pour plus d’informations sur la validation des jetons d’identité d’utilisateur Exchange, reportez-vous à l’article [Valider un jeton d’identité Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="9f021-117">For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span>

<span data-ttu-id="9f021-118">Une fois validée et décodée, la charge utile du jeton ressemble à ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="9f021-118">Once validated and decoded, the payload of the token looks something like the following.</span></span>

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a><span data-ttu-id="9f021-119">Mapper le jeton à un utilisateur dans votre back end
</span><span class="sxs-lookup"><span data-stu-id="9f021-119">Map the token to a user in your backend</span></span>

<span data-ttu-id="9f021-120">Votre service principal peut calculer un ID d’utilisateur unique à partir du jeton et le mapper à un utilisateur dans votre système d’utilisateur interne.</span><span class="sxs-lookup"><span data-stu-id="9f021-120">Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system.</span></span> <span data-ttu-id="9f021-121">Par exemple, si vous utilisez une base de données pour stocker des utilisateurs, vous pouvez ajouter cet ID unique à l’enregistrement de l’utilisateur dans votre base de données.</span><span class="sxs-lookup"><span data-stu-id="9f021-121">For example, if you use a database to store users, you could add this unique ID to the user's record in your database.</span></span>

### <a name="generate-a-unique-id"></a><span data-ttu-id="9f021-122">Génération d’un ID unique</span><span class="sxs-lookup"><span data-stu-id="9f021-122">Generate a unique ID</span></span>

<span data-ttu-id="9f021-123">Nous vous recommandons d’utiliser une combinaison des propriétés `msexchuid` et `amurl`.</span><span class="sxs-lookup"><span data-stu-id="9f021-123">We recommend that you use a combination of the `msexchuid` and `amurl` properties.</span></span> <span data-ttu-id="9f021-124">Par exemple, vous pouvez concaténer les deux valeurs et générer une chaîne codée au format base64.</span><span class="sxs-lookup"><span data-stu-id="9f021-124">For example, you could concatenate the two values together and generate a base 64-encoded string.</span></span> <span data-ttu-id="9f021-125">Cette valeur peut être générée en toute fiabilité à partir du jeton à chaque fois. Ainsi, vous pouvez mapper un jeton d’identité d’utilisateur Exchange à l’utilisateur dans votre système.</span><span class="sxs-lookup"><span data-stu-id="9f021-125">This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.</span></span>

### <a name="check-the-user"></a><span data-ttu-id="9f021-126">Vérification de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="9f021-126">Check the user</span></span>

<span data-ttu-id="9f021-127">Avec l’ID unique généré, l’étape suivante consiste à vérifier la présence d’un utilisateur dans votre système avec cet ID associé.</span><span class="sxs-lookup"><span data-stu-id="9f021-127">With the unique ID generated, the next step is to check for a user in your system with that associated ID.</span></span>

- <span data-ttu-id="9f021-128">Si vous trouvez l’utilisateur, le back end considère la requête comme authentifiée et autorise sa poursuite.
</span><span class="sxs-lookup"><span data-stu-id="9f021-128">If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.</span></span>

- <span data-ttu-id="9f021-129">Si l’utilisateur est introuvable, le back end renvoie une erreur indiquant que l’utilisateur doit se connecter. 
</span><span class="sxs-lookup"><span data-stu-id="9f021-129">If the user is not found, then the back-end returns an error indicating that the user needs to sign in.</span></span> <span data-ttu-id="9f021-130">Le complément invite ensuite l’utilisateur à se connecter au back end à l’aide de votre méthode d’authentification.
</span><span class="sxs-lookup"><span data-stu-id="9f021-130">The add-in then prompts the user to sign in to the back-end using your existing authentication method.</span></span> <span data-ttu-id="9f021-131">Une fois l’utilisateur authentifié, le jeton d’identité d’utilisateur Exchange est envoyé avec les détails de l’authentification utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9f021-131">Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details.</span></span> <span data-ttu-id="9f021-132">Le back end peut ensuite mettre à jour l’enregistrement de l’utilisateur dans votre système avec l’ID unique.
</span><span class="sxs-lookup"><span data-stu-id="9f021-132">The back-end can then update the user's record in your system with the unique ID.</span></span>