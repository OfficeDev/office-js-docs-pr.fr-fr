---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450687"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="9c94c-102">SourceLocation, élément</span><span class="sxs-lookup"><span data-stu-id="9c94c-102">SourceLocation element</span></span>

<span data-ttu-id="9c94c-103">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="9c94c-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="9c94c-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="9c94c-104">Attributes</span></span>

| <span data-ttu-id="9c94c-105">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="9c94c-105">**Attribute**</span></span> | <span data-ttu-id="9c94c-106">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="9c94c-106">**Required**</span></span> | <span data-ttu-id="9c94c-107">**Description**</span><span class="sxs-lookup"><span data-stu-id="9c94c-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="9c94c-108">resid</span><span class="sxs-lookup"><span data-stu-id="9c94c-108">resid</span></span>         | <span data-ttu-id="9c94c-109">Oui</span><span class="sxs-lookup"><span data-stu-id="9c94c-109">Yes</span></span>          | <span data-ttu-id="9c94c-110">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="9c94c-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="9c94c-111">Child, élément</span><span class="sxs-lookup"><span data-stu-id="9c94c-111">Child elements</span></span>

<span data-ttu-id="9c94c-112">Aucun</span><span class="sxs-lookup"><span data-stu-id="9c94c-112">None</span></span>

## <a name="example"></a><span data-ttu-id="9c94c-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="9c94c-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
