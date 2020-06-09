---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611763"
---
# <a name="metadata-element"></a><span data-ttu-id="14d35-103">Élément de métadonnées</span><span class="sxs-lookup"><span data-stu-id="14d35-103">Metadata element</span></span>

<span data-ttu-id="14d35-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="14d35-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="14d35-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="14d35-105">Attributes</span></span>

<span data-ttu-id="14d35-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="14d35-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="14d35-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="14d35-107">Child elements</span></span>

|  <span data-ttu-id="14d35-108">Élément</span><span class="sxs-lookup"><span data-stu-id="14d35-108">Element</span></span>  |  <span data-ttu-id="14d35-109">Requis</span><span class="sxs-lookup"><span data-stu-id="14d35-109">Required</span></span>  |  <span data-ttu-id="14d35-110">Description</span><span class="sxs-lookup"><span data-stu-id="14d35-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="14d35-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="14d35-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="14d35-112">Oui</span><span class="sxs-lookup"><span data-stu-id="14d35-112">Yes</span></span>  | <span data-ttu-id="14d35-113">Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="14d35-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="14d35-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="14d35-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
