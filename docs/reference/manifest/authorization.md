---
title: Authorization, élément dans le fichier manifeste
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477942"
---
# <a name="authorization-element"></a><span data-ttu-id="d19d3-102">Authorization, élément</span><span class="sxs-lookup"><span data-stu-id="d19d3-102">Authorization element</span></span>

<span data-ttu-id="d19d3-103">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="d19d3-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="d19d3-104">**Authorization** est un élément enfant de [](authorizations.md) l’élément Authorizations dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="d19d3-104">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d19d3-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d19d3-105">Child elements</span></span>

|  <span data-ttu-id="d19d3-106">Élément</span><span class="sxs-lookup"><span data-stu-id="d19d3-106">Element</span></span> |  <span data-ttu-id="d19d3-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d19d3-107">Required</span></span>  |  <span data-ttu-id="d19d3-108">Description</span><span class="sxs-lookup"><span data-stu-id="d19d3-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d19d3-109">**Resource**</span><span class="sxs-lookup"><span data-stu-id="d19d3-109">**Resource**</span></span>  |  <span data-ttu-id="d19d3-110">Oui</span><span class="sxs-lookup"><span data-stu-id="d19d3-110">Yes</span></span>   |  <span data-ttu-id="d19d3-111">Spécifie l’URL de la ressource externe.</span><span class="sxs-lookup"><span data-stu-id="d19d3-111">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="d19d3-112">Scopes</span><span class="sxs-lookup"><span data-stu-id="d19d3-112">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="d19d3-113">Oui</span><span class="sxs-lookup"><span data-stu-id="d19d3-113">Yes</span></span>  |  <span data-ttu-id="d19d3-114">Spécifie les autorisations dont le complément a besoin pour la ressource.</span><span class="sxs-lookup"><span data-stu-id="d19d3-114">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="d19d3-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="d19d3-115">Example</span></span>

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
