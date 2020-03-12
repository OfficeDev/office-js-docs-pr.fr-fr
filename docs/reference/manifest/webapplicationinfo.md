---
title: Élément WebApplicationInfo dans le fichier manifeste
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b6cf82776f683929845df83c642b28ad024d665a
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596731"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="68a9c-102">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="68a9c-102">WebApplicationInfo element</span></span>

<span data-ttu-id="68a9c-103">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="68a9c-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="68a9c-104">En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.</span><span class="sxs-lookup"><span data-stu-id="68a9c-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="68a9c-105">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="68a9c-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="68a9c-106">L’API d’authentification unique est actuellement prise en charge en préversion pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="68a9c-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="68a9c-107">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="68a9c-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="68a9c-108">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="68a9c-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="68a9c-109">Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="68a9c-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="68a9c-110">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="68a9c-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="68a9c-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68a9c-111">Child elements</span></span>

|  <span data-ttu-id="68a9c-112">Élément</span><span class="sxs-lookup"><span data-stu-id="68a9c-112">Element</span></span> |  <span data-ttu-id="68a9c-113">Requis</span><span class="sxs-lookup"><span data-stu-id="68a9c-113">Required</span></span>  |  <span data-ttu-id="68a9c-114">Description</span><span class="sxs-lookup"><span data-stu-id="68a9c-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="68a9c-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="68a9c-115">**Id**</span></span>    |  <span data-ttu-id="68a9c-116">Oui</span><span class="sxs-lookup"><span data-stu-id="68a9c-116">Yes</span></span>   |  <span data-ttu-id="68a9c-117">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="68a9c-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="68a9c-118">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="68a9c-118">**MsaId**</span></span>    |  <span data-ttu-id="68a9c-119">Non</span><span class="sxs-lookup"><span data-stu-id="68a9c-119">No</span></span>   |  <span data-ttu-id="68a9c-120">ID client de l’application Web de votre complément pour MSA, tel qu’inscrit dans msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="68a9c-120">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="68a9c-121">**Resource**</span><span class="sxs-lookup"><span data-stu-id="68a9c-121">**Resource**</span></span>  |  <span data-ttu-id="68a9c-122">Oui</span><span class="sxs-lookup"><span data-stu-id="68a9c-122">Yes</span></span>   |  <span data-ttu-id="68a9c-123">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="68a9c-123">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="68a9c-124">Scopes</span><span class="sxs-lookup"><span data-stu-id="68a9c-124">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="68a9c-125">Oui</span><span class="sxs-lookup"><span data-stu-id="68a9c-125">Yes</span></span>  |  <span data-ttu-id="68a9c-126">Spécifie les autorisations dont le complément a besoin pour une ressource, telle que Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="68a9c-126">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="68a9c-127">Autorisations</span><span class="sxs-lookup"><span data-stu-id="68a9c-127">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="68a9c-128">Non</span><span class="sxs-lookup"><span data-stu-id="68a9c-128">No</span></span>   | <span data-ttu-id="68a9c-129">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="68a9c-129">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="68a9c-130">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="68a9c-130">WebApplicationInfo example</span></span>

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
