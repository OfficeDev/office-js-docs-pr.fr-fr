---
title: Authorization, élément dans le fichier manifeste
description: Spécifie une ressource externe à laquelle l’application Web du complément doit disposer d’une autorisation et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294261"
---
# <a name="authorization-element"></a><span data-ttu-id="9fb1a-103">Authorization, élément</span><span class="sxs-lookup"><span data-stu-id="9fb1a-103">Authorization element</span></span>

<span data-ttu-id="9fb1a-104">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="9fb1a-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="9fb1a-105">**Authorization** est un élément enfant de l’élément [Authorizations](authorizations.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="9fb1a-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9fb1a-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9fb1a-106">Child elements</span></span>

|  <span data-ttu-id="9fb1a-107">Élément</span><span class="sxs-lookup"><span data-stu-id="9fb1a-107">Element</span></span> |  <span data-ttu-id="9fb1a-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9fb1a-108">Required</span></span>  |  <span data-ttu-id="9fb1a-109">Description</span><span class="sxs-lookup"><span data-stu-id="9fb1a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9fb1a-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="9fb1a-110">**Resource**</span></span>  |  <span data-ttu-id="9fb1a-111">Oui</span><span class="sxs-lookup"><span data-stu-id="9fb1a-111">Yes</span></span>   |  <span data-ttu-id="9fb1a-112">Spécifie l’URL de la ressource externe.</span><span class="sxs-lookup"><span data-stu-id="9fb1a-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="9fb1a-113">Scopes</span><span class="sxs-lookup"><span data-stu-id="9fb1a-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="9fb1a-114">Oui</span><span class="sxs-lookup"><span data-stu-id="9fb1a-114">Yes</span></span>  |  <span data-ttu-id="9fb1a-115">Spécifie les autorisations dont le complément a besoin pour la ressource.</span><span class="sxs-lookup"><span data-stu-id="9fb1a-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="9fb1a-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="9fb1a-116">Example</span></span>

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
