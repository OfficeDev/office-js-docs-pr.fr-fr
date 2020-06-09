---
title: Élément Scopes dans le fichier manifeste
description: L’élément Scopes contient les autorisations dont le complément a besoin pour se connecter à une ressource externe.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612240"
---
# <a name="scopes-element"></a><span data-ttu-id="b455d-103">Élément Scopes</span><span class="sxs-lookup"><span data-stu-id="b455d-103">Scopes element</span></span>

<span data-ttu-id="b455d-104">Contient les autorisations dont le complément a besoin pour une ressource externe, telle que Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b455d-104">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="b455d-105">Lorsque Microsoft Graph est la ressource, AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement.</span><span class="sxs-lookup"><span data-stu-id="b455d-105">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="b455d-106">Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b455d-106">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="b455d-107">**Scopes** est un élément enfant des éléments [WebApplicationInfo](webapplicationinfo.md) et [authorization](authorization.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="b455d-107">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b455d-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b455d-108">Child elements</span></span>

|  <span data-ttu-id="b455d-109">Élément</span><span class="sxs-lookup"><span data-stu-id="b455d-109">Element</span></span> |  <span data-ttu-id="b455d-110">Requis</span><span class="sxs-lookup"><span data-stu-id="b455d-110">Required</span></span>  |  <span data-ttu-id="b455d-111">Description</span><span class="sxs-lookup"><span data-stu-id="b455d-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b455d-112">**Scope**</span><span class="sxs-lookup"><span data-stu-id="b455d-112">**Scope**</span></span>                |  <span data-ttu-id="b455d-113">Oui</span><span class="sxs-lookup"><span data-stu-id="b455d-113">Yes</span></span>     |   <span data-ttu-id="b455d-114">Nom d’une autorisation ; par exemple, files. Read. All ou Profile.</span><span class="sxs-lookup"><span data-stu-id="b455d-114">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="b455d-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="b455d-115">Example</span></span>

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
