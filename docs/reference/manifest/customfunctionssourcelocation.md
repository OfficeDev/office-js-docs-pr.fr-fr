---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641381"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="1b30f-103">Élément SourceLocation (fonctions personnalisées)</span><span class="sxs-lookup"><span data-stu-id="1b30f-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="1b30f-104">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1b30f-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="1b30f-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="1b30f-105">Attributes</span></span>

| <span data-ttu-id="1b30f-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="1b30f-106">Attribute</span></span> | <span data-ttu-id="1b30f-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1b30f-107">Required</span></span> | <span data-ttu-id="1b30f-108">Description</span><span class="sxs-lookup"><span data-stu-id="1b30f-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="1b30f-109">resid</span><span class="sxs-lookup"><span data-stu-id="1b30f-109">resid</span></span>     | <span data-ttu-id="1b30f-110">Oui</span><span class="sxs-lookup"><span data-stu-id="1b30f-110">Yes</span></span>      | <span data-ttu-id="1b30f-111">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="1b30f-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="1b30f-112">Child, élément</span><span class="sxs-lookup"><span data-stu-id="1b30f-112">Child elements</span></span>

<span data-ttu-id="1b30f-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="1b30f-113">None</span></span>

## <a name="example"></a><span data-ttu-id="1b30f-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="1b30f-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
