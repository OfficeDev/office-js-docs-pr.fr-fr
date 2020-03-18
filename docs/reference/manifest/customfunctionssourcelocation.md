---
title: Élément SourceLocation dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720686"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="35c46-103">SourceLocation, élément</span><span class="sxs-lookup"><span data-stu-id="35c46-103">SourceLocation element</span></span>

<span data-ttu-id="35c46-104">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="35c46-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="35c46-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="35c46-105">Attributes</span></span>

| <span data-ttu-id="35c46-106">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="35c46-106">**Attribute**</span></span> | <span data-ttu-id="35c46-107">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="35c46-107">**Required**</span></span> | <span data-ttu-id="35c46-108">**Description**</span><span class="sxs-lookup"><span data-stu-id="35c46-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="35c46-109">resid</span><span class="sxs-lookup"><span data-stu-id="35c46-109">resid</span></span>         | <span data-ttu-id="35c46-110">Oui</span><span class="sxs-lookup"><span data-stu-id="35c46-110">Yes</span></span>          | <span data-ttu-id="35c46-111">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="35c46-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="35c46-112">Child, élément</span><span class="sxs-lookup"><span data-stu-id="35c46-112">Child elements</span></span>

<span data-ttu-id="35c46-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="35c46-113">None</span></span>

## <a name="example"></a><span data-ttu-id="35c46-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="35c46-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
