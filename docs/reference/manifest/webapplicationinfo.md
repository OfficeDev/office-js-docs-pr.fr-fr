---
title: Élément WebApplicationInfo dans le fichier manifeste
description: Documentation de référence de l’élément WebApplicationInfo pour les fichiers manifeste (XML) des applications Office.
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: abbb4b97047fda378da71963f3f522fae4d72ccc
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839704"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="c3733-103">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="c3733-103">WebApplicationInfo element</span></span>

<span data-ttu-id="c3733-104">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="c3733-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="c3733-105">Ressource OAuth 2.0  pour laquelle l’application cliente Office peut avoir besoin d’autorisations.</span><span class="sxs-lookup"><span data-stu-id="c3733-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="c3733-106">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c3733-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="c3733-107">L’API d' sign-on unique est actuellement prise en charge pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c3733-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="c3733-108">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c3733-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="c3733-109">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="c3733-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="c3733-110">Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="c3733-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="c3733-111">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="c3733-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="c3733-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c3733-112">Child elements</span></span>

|  <span data-ttu-id="c3733-113">Élément</span><span class="sxs-lookup"><span data-stu-id="c3733-113">Element</span></span> |  <span data-ttu-id="c3733-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="c3733-114">Required</span></span>  |  <span data-ttu-id="c3733-115">Description</span><span class="sxs-lookup"><span data-stu-id="c3733-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c3733-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="c3733-116">**Id**</span></span>    |  <span data-ttu-id="c3733-117">Oui</span><span class="sxs-lookup"><span data-stu-id="c3733-117">Yes</span></span>   |  <span data-ttu-id="c3733-118">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="c3733-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="c3733-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="c3733-119">**MsaId**</span></span>    |  <span data-ttu-id="c3733-120">Non</span><span class="sxs-lookup"><span data-stu-id="c3733-120">No</span></span>   |  <span data-ttu-id="c3733-121">ID client de l’application web de votre add-in pour MSA tel qu’inscrit dans msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="c3733-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="c3733-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="c3733-122">**Resource**</span></span>  |  <span data-ttu-id="c3733-123">Oui</span><span class="sxs-lookup"><span data-stu-id="c3733-123">Yes</span></span>   |  <span data-ttu-id="c3733-124">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="c3733-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="c3733-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="c3733-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="c3733-126">Oui</span><span class="sxs-lookup"><span data-stu-id="c3733-126">Yes</span></span>  |  <span data-ttu-id="c3733-127">Spécifie les autorisations dont le add-in a besoin pour une ressource, telle que Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c3733-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="c3733-128">Autorisations</span><span class="sxs-lookup"><span data-stu-id="c3733-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="c3733-129">Non</span><span class="sxs-lookup"><span data-stu-id="c3733-129">No</span></span>   | <span data-ttu-id="c3733-130">Spécifie les ressources externes à qui l’application web du add-in a besoin d’autorisation et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="c3733-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="c3733-131">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="c3733-131">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```