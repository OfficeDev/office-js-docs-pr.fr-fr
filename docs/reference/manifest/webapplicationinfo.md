---
title: Élément WebApplicationInfo dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ae1f2ea43e795d8f4ac2f634785fc69b0ca7a3d0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457627"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="b6b7c-102">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="b6b7c-102">WebApplicationInfo element</span></span>

<span data-ttu-id="b6b7c-103">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="b6b7c-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="b6b7c-104">En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="b6b7c-105">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="b6b7c-106">L’API d’authentification unique est actuellement prise en charge en préversion pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="b6b7c-107">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, voir [Ensembles de conditions requises de l’API d’identité](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b6b7c-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="b6b7c-108">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="b6b7c-109">Pour savoir comment procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="b6b7c-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="b6b7c-110">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="b6b7c-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b6b7c-111">Child elements</span></span>

|  <span data-ttu-id="b6b7c-112">Élément</span><span class="sxs-lookup"><span data-stu-id="b6b7c-112">Element</span></span> |  <span data-ttu-id="b6b7c-113">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b6b7c-113">Required</span></span>  |  <span data-ttu-id="b6b7c-114">Description</span><span class="sxs-lookup"><span data-stu-id="b6b7c-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b6b7c-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="b6b7c-115">**Id**</span></span>    |  <span data-ttu-id="b6b7c-116">Oui</span><span class="sxs-lookup"><span data-stu-id="b6b7c-116">Yes</span></span>   |  <span data-ttu-id="b6b7c-117">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="b6b7c-118">**Resource**</span><span class="sxs-lookup"><span data-stu-id="b6b7c-118">**Resource**</span></span>  |  <span data-ttu-id="b6b7c-119">Oui</span><span class="sxs-lookup"><span data-stu-id="b6b7c-119">Yes</span></span>   |  <span data-ttu-id="b6b7c-120">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="b6b7c-121">Scopes</span><span class="sxs-lookup"><span data-stu-id="b6b7c-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="b6b7c-122">Non</span><span class="sxs-lookup"><span data-stu-id="b6b7c-122">No</span></span>  |  <span data-ttu-id="b6b7c-123">Spécifie les autorisations dont le complément a besoin pour Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="b6b7c-124">À l’heure actuelle, il est nécessaire que les ressources de votre complément correspondent à son hôte.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="b6b7c-125">Office ne demandera pas un jeton pour un complément à moins de pouvoir prouver qu’il en est le propriétaire ; à l’heure actuelle, ceci s’effectue en hébergeant le complément sous le nom de domaine complet de la ressource.</span><span class="sxs-lookup"><span data-stu-id="b6b7c-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="b6b7c-126">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="b6b7c-126">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
