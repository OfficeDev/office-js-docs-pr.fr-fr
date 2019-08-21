---
title: Élément Scopes dans le fichier manifeste
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1e36bdcd0cdcaa8c842e924c2543d56bdc4e26a7
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477732"
---
# <a name="scopes-element"></a><span data-ttu-id="1ac45-102">Élément Scopes</span><span class="sxs-lookup"><span data-stu-id="1ac45-102">Scopes element</span></span>

<span data-ttu-id="1ac45-103">Contient les autorisations dont le complément a besoin pour une ressource externe, telle que Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1ac45-103">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="1ac45-104">Lorsque Microsoft Graph est la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement.</span><span class="sxs-lookup"><span data-stu-id="1ac45-104">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="1ac45-105">Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1ac45-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="1ac45-106">**Scopes** est un élément enfant des éléments [WebApplicationInfo](webapplicationinfo.md) et [authorization](authorization.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="1ac45-106">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1ac45-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1ac45-107">Child elements</span></span>

|  <span data-ttu-id="1ac45-108">Élément</span><span class="sxs-lookup"><span data-stu-id="1ac45-108">Element</span></span> |  <span data-ttu-id="1ac45-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1ac45-109">Required</span></span>  |  <span data-ttu-id="1ac45-110">Description</span><span class="sxs-lookup"><span data-stu-id="1ac45-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1ac45-111">**Scope**</span><span class="sxs-lookup"><span data-stu-id="1ac45-111">**Scope**</span></span>                |  <span data-ttu-id="1ac45-112">Oui</span><span class="sxs-lookup"><span data-stu-id="1ac45-112">Yes</span></span>     |   <span data-ttu-id="1ac45-113">Nom d’une autorisation; par exemple, files. Read. All ou Profile.</span><span class="sxs-lookup"><span data-stu-id="1ac45-113">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="1ac45-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ac45-114">Example</span></span>

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
