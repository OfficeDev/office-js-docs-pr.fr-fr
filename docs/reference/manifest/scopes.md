---
title: Élément Scopes dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771785"
---
# <a name="scopes-element"></a><span data-ttu-id="848a2-102">Élément Scopes</span><span class="sxs-lookup"><span data-stu-id="848a2-102">Scopes element</span></span>

<span data-ttu-id="848a2-103">Contient des autorisations Microsoft Graph requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="848a2-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="848a2-104">AppSource utilise l’élément Scopes pour créer une boîte de dialogue de consentement.</span><span class="sxs-lookup"><span data-stu-id="848a2-104">AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="848a2-105">Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="848a2-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="848a2-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="848a2-106">Child elements</span></span>

|  <span data-ttu-id="848a2-107">Élément</span><span class="sxs-lookup"><span data-stu-id="848a2-107">Element</span></span> |  <span data-ttu-id="848a2-108">Type</span><span class="sxs-lookup"><span data-stu-id="848a2-108">Type</span></span>  |  <span data-ttu-id="848a2-109">Description</span><span class="sxs-lookup"><span data-stu-id="848a2-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="848a2-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="848a2-110">**Scope**</span></span>                |  <span data-ttu-id="848a2-111">string</span><span class="sxs-lookup"><span data-stu-id="848a2-111">string</span></span>     |   <span data-ttu-id="848a2-112">Nom d’une autorisation Microsoft Graph ; par exemple, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="848a2-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="848a2-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="848a2-113">Example</span></span>

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
