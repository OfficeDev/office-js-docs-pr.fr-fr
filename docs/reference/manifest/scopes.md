---
title: Élément Scopes dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432640"
---
# <a name="scopes-element"></a><span data-ttu-id="649b7-102">Élément Scopes</span><span class="sxs-lookup"><span data-stu-id="649b7-102">Scopes element</span></span>

<span data-ttu-id="649b7-103">Contient des autorisations Microsoft Graph requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="649b7-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="649b7-104">L’Office Store se sert de l’élément Scope pour créer une boîte de dialogue de consentement.</span><span class="sxs-lookup"><span data-stu-id="649b7-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="649b7-105">Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="649b7-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="649b7-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="649b7-106">Child elements</span></span>

|  <span data-ttu-id="649b7-107">Élément</span><span class="sxs-lookup"><span data-stu-id="649b7-107">Element</span></span> |  <span data-ttu-id="649b7-108">Type</span><span class="sxs-lookup"><span data-stu-id="649b7-108">Type</span></span>  |  <span data-ttu-id="649b7-109">Description</span><span class="sxs-lookup"><span data-stu-id="649b7-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="649b7-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="649b7-110">**Scope**</span></span>                |  <span data-ttu-id="649b7-111">string</span><span class="sxs-lookup"><span data-stu-id="649b7-111">string</span></span>     |   <span data-ttu-id="649b7-112">Nom d’une autorisation Microsoft Graph ; par exemple, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="649b7-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="649b7-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="649b7-113">Example</span></span>

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
