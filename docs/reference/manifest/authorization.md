---
title: Authorization, élément dans le fichier manifeste
description: Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8c6249706b8eef11f579378fe5c9dc83016d17c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608760"
---
# <a name="authorization-element"></a><span data-ttu-id="00b16-103">Authorization, élément</span><span class="sxs-lookup"><span data-stu-id="00b16-103">Authorization element</span></span>

<span data-ttu-id="00b16-104">Spécifie les ressources externes auxquelles l’application Web du complément doit disposer et les autorisations requises.</span><span class="sxs-lookup"><span data-stu-id="00b16-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="00b16-105">**Authorization** est un élément enfant de l’élément [Authorizations](authorizations.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="00b16-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="00b16-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="00b16-106">Child elements</span></span>

|  <span data-ttu-id="00b16-107">Élément</span><span class="sxs-lookup"><span data-stu-id="00b16-107">Element</span></span> |  <span data-ttu-id="00b16-108">Requis</span><span class="sxs-lookup"><span data-stu-id="00b16-108">Required</span></span>  |  <span data-ttu-id="00b16-109">Description</span><span class="sxs-lookup"><span data-stu-id="00b16-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="00b16-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="00b16-110">**Resource**</span></span>  |  <span data-ttu-id="00b16-111">Oui</span><span class="sxs-lookup"><span data-stu-id="00b16-111">Yes</span></span>   |  <span data-ttu-id="00b16-112">Spécifie l’URL de la ressource externe.</span><span class="sxs-lookup"><span data-stu-id="00b16-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="00b16-113">Scopes</span><span class="sxs-lookup"><span data-stu-id="00b16-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="00b16-114">Oui</span><span class="sxs-lookup"><span data-stu-id="00b16-114">Yes</span></span>  |  <span data-ttu-id="00b16-115">Spécifie les autorisations dont le complément a besoin pour la ressource.</span><span class="sxs-lookup"><span data-stu-id="00b16-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="00b16-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="00b16-116">Example</span></span>

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
