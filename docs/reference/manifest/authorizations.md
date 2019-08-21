---
title: Authorization, élément dans le fichier manifeste
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6a271423ddd549431c2f580e2793faab3c49090e
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477956"
---
# <a name="authorizations-element"></a><span data-ttu-id="39e1a-102">Authorizations, élément</span><span class="sxs-lookup"><span data-stu-id="39e1a-102">Authorizations element</span></span>

<span data-ttu-id="39e1a-103">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="39e1a-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="39e1a-104">**Authorizations** est un élément enfant de l’élément [WebApplicationInfo](webapplicationinfo.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="39e1a-104">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="39e1a-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="39e1a-105">Child elements</span></span>

|  <span data-ttu-id="39e1a-106">Élément</span><span class="sxs-lookup"><span data-stu-id="39e1a-106">Element</span></span> |  <span data-ttu-id="39e1a-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="39e1a-107">Required</span></span>  |  <span data-ttu-id="39e1a-108">Description</span><span class="sxs-lookup"><span data-stu-id="39e1a-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="39e1a-109">Autorisation</span><span class="sxs-lookup"><span data-stu-id="39e1a-109">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="39e1a-110">Oui</span><span class="sxs-lookup"><span data-stu-id="39e1a-110">Yes</span></span>     |   <span data-ttu-id="39e1a-111">Identifie une ressource externe dont l’application Web du complément a besoin d’autorisation, ainsi que les étendues (autorisations) dont elle a besoin.</span><span class="sxs-lookup"><span data-stu-id="39e1a-111">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="39e1a-112">Exemple</span><span class="sxs-lookup"><span data-stu-id="39e1a-112">Example</span></span>

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
