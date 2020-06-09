---
title: Élément SourceLocation dans le fichier manifeste
description: Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612311"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="49a4b-103">SourceLocation, élément</span><span class="sxs-lookup"><span data-stu-id="49a4b-103">SourceLocation element</span></span>

<span data-ttu-id="49a4b-104">Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="49a4b-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="49a4b-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="49a4b-105">Attributes</span></span>

| <span data-ttu-id="49a4b-106">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="49a4b-106">**Attribute**</span></span> | <span data-ttu-id="49a4b-107">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="49a4b-107">**Required**</span></span> | <span data-ttu-id="49a4b-108">**Description**</span><span class="sxs-lookup"><span data-stu-id="49a4b-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="49a4b-109">resid</span><span class="sxs-lookup"><span data-stu-id="49a4b-109">resid</span></span>         | <span data-ttu-id="49a4b-110">Oui</span><span class="sxs-lookup"><span data-stu-id="49a4b-110">Yes</span></span>          | <span data-ttu-id="49a4b-111">Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste.</span><span class="sxs-lookup"><span data-stu-id="49a4b-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="49a4b-112">Child, élément</span><span class="sxs-lookup"><span data-stu-id="49a4b-112">Child elements</span></span>

<span data-ttu-id="49a4b-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="49a4b-113">None</span></span>

## <a name="example"></a><span data-ttu-id="49a4b-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="49a4b-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
