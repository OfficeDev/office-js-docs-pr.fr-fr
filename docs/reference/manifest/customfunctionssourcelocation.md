---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432405"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="ac81a-102">SourceLocation, élément</span><span class="sxs-lookup"><span data-stu-id="ac81a-102">SourceLocation element</span></span>

<span data-ttu-id="ac81a-103">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ac81a-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="ac81a-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="ac81a-104">Attributes</span></span>

| <span data-ttu-id="ac81a-105">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="ac81a-105">**Attribute**</span></span> | <span data-ttu-id="ac81a-106">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="ac81a-106">**Required**</span></span> | <span data-ttu-id="ac81a-107">**Description**</span><span class="sxs-lookup"><span data-stu-id="ac81a-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="ac81a-108">resid</span><span class="sxs-lookup"><span data-stu-id="ac81a-108">resid</span></span>         | <span data-ttu-id="ac81a-109">Oui</span><span class="sxs-lookup"><span data-stu-id="ac81a-109">Yes</span></span>          | <span data-ttu-id="ac81a-110">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="ac81a-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="ac81a-111">Child, élément</span><span class="sxs-lookup"><span data-stu-id="ac81a-111">Child elements</span></span>

<span data-ttu-id="ac81a-112">Aucun</span><span class="sxs-lookup"><span data-stu-id="ac81a-112">None</span></span>

## <a name="example"></a><span data-ttu-id="ac81a-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac81a-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```