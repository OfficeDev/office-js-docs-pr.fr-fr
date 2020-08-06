---
title: Élément WebApplicationInfo dans le fichier manifeste
description: Documentation de référence de l’élément VersionOverrides pour les fichiers manifeste des compléments Office (XML).
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: b64739f388720abb872fe2407f7e7d3031b5ec70
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573118"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="07381-103">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="07381-103">WebApplicationInfo element</span></span>

<span data-ttu-id="07381-104">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="07381-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="07381-105">En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.</span><span class="sxs-lookup"><span data-stu-id="07381-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="07381-106">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="07381-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="07381-107">L’API d’authentification unique est actuellement prise en charge pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="07381-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="07381-108">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="07381-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="07381-109">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="07381-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="07381-110">Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="07381-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="07381-111">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="07381-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="07381-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="07381-112">Child elements</span></span>

|  <span data-ttu-id="07381-113">Élément</span><span class="sxs-lookup"><span data-stu-id="07381-113">Element</span></span> |  <span data-ttu-id="07381-114">Requis</span><span class="sxs-lookup"><span data-stu-id="07381-114">Required</span></span>  |  <span data-ttu-id="07381-115">Description</span><span class="sxs-lookup"><span data-stu-id="07381-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="07381-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="07381-116">**Id**</span></span>    |  <span data-ttu-id="07381-117">Oui</span><span class="sxs-lookup"><span data-stu-id="07381-117">Yes</span></span>   |  <span data-ttu-id="07381-118">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="07381-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="07381-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="07381-119">**MsaId**</span></span>    |  <span data-ttu-id="07381-120">Non</span><span class="sxs-lookup"><span data-stu-id="07381-120">No</span></span>   |  <span data-ttu-id="07381-121">ID client de l’application Web de votre complément pour MSA, tel qu’inscrit dans msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="07381-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="07381-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="07381-122">**Resource**</span></span>  |  <span data-ttu-id="07381-123">Oui</span><span class="sxs-lookup"><span data-stu-id="07381-123">Yes</span></span>   |  <span data-ttu-id="07381-124">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="07381-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="07381-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="07381-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="07381-126">Oui</span><span class="sxs-lookup"><span data-stu-id="07381-126">Yes</span></span>  |  <span data-ttu-id="07381-127">Spécifie les autorisations dont le complément a besoin pour une ressource, telle que Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="07381-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="07381-128">Autorisations</span><span class="sxs-lookup"><span data-stu-id="07381-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="07381-129">Non</span><span class="sxs-lookup"><span data-stu-id="07381-129">No</span></span>   | <span data-ttu-id="07381-130">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="07381-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="07381-131">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="07381-131">WebApplicationInfo example</span></span>

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
