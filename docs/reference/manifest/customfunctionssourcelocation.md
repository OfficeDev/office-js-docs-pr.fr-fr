---
title: Élément SourceLocation pour les fonctions personnalisées dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771381"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="b0177-103">Élément SourceLocation (fonctions personnalisées)</span><span class="sxs-lookup"><span data-stu-id="b0177-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="b0177-104">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b0177-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="b0177-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="b0177-105">Attributes</span></span>

| <span data-ttu-id="b0177-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="b0177-106">Attribute</span></span> | <span data-ttu-id="b0177-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b0177-107">Required</span></span> | <span data-ttu-id="b0177-108">Description</span><span class="sxs-lookup"><span data-stu-id="b0177-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="b0177-109">resid</span><span class="sxs-lookup"><span data-stu-id="b0177-109">resid</span></span>     | <span data-ttu-id="b0177-110">Oui</span><span class="sxs-lookup"><span data-stu-id="b0177-110">Yes</span></span>      | <span data-ttu-id="b0177-111">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="b0177-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> <span data-ttu-id="b0177-112">Il ne peut pas comporter plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="b0177-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="b0177-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b0177-113">Child elements</span></span>

<span data-ttu-id="b0177-114">Aucun</span><span class="sxs-lookup"><span data-stu-id="b0177-114">None</span></span>

## <a name="example"></a><span data-ttu-id="b0177-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="b0177-115">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
