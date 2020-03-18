---
title: Élément de métadonnées dans le fichier manifest
description: L’élément Metadata définit les paramètres de métadonnées qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718068"
---
# <a name="metadata-element"></a><span data-ttu-id="afc37-103">Élément de métadonnées</span><span class="sxs-lookup"><span data-stu-id="afc37-103">Metadata element</span></span>

<span data-ttu-id="afc37-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="afc37-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="afc37-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="afc37-105">Attributes</span></span>

<span data-ttu-id="afc37-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="afc37-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="afc37-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="afc37-107">Child elements</span></span>

|  <span data-ttu-id="afc37-108">Élément</span><span class="sxs-lookup"><span data-stu-id="afc37-108">Element</span></span>  |  <span data-ttu-id="afc37-109">Requis</span><span class="sxs-lookup"><span data-stu-id="afc37-109">Required</span></span>  |  <span data-ttu-id="afc37-110">Description</span><span class="sxs-lookup"><span data-stu-id="afc37-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="afc37-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="afc37-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="afc37-112">Oui</span><span class="sxs-lookup"><span data-stu-id="afc37-112">Yes</span></span>  | <span data-ttu-id="afc37-113">Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="afc37-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="afc37-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="afc37-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
