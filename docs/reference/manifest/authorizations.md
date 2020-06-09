---
title: Authorization, élément dans le fichier manifeste
description: Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 675585f99fc6261a2145219d553f02b9f9abded3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608753"
---
# <a name="authorizations-element"></a><span data-ttu-id="cb631-103">Authorizations, élément</span><span class="sxs-lookup"><span data-stu-id="cb631-103">Authorizations element</span></span>

<span data-ttu-id="cb631-104">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="cb631-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="cb631-105">**Authorizations** est un élément enfant de l’élément [WebApplicationInfo](webapplicationinfo.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="cb631-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cb631-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="cb631-106">Child elements</span></span>

|  <span data-ttu-id="cb631-107">Élément</span><span class="sxs-lookup"><span data-stu-id="cb631-107">Element</span></span> |  <span data-ttu-id="cb631-108">Requis</span><span class="sxs-lookup"><span data-stu-id="cb631-108">Required</span></span>  |  <span data-ttu-id="cb631-109">Description</span><span class="sxs-lookup"><span data-stu-id="cb631-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cb631-110">Autorisation</span><span class="sxs-lookup"><span data-stu-id="cb631-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="cb631-111">Oui</span><span class="sxs-lookup"><span data-stu-id="cb631-111">Yes</span></span>     |   <span data-ttu-id="cb631-112">Identifie une ressource externe dont l’application Web du complément a besoin d’autorisation, ainsi que les étendues (autorisations) dont elle a besoin.</span><span class="sxs-lookup"><span data-stu-id="cb631-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="cb631-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="cb631-113">Example</span></span>

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
