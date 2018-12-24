---
title: Élément de métadonnées dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432661"
---
# <a name="metadata-element"></a><span data-ttu-id="441fc-102">Élément de métadonnées</span><span class="sxs-lookup"><span data-stu-id="441fc-102">MetaData element</span></span>

<span data-ttu-id="441fc-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="441fc-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="441fc-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="441fc-104">Attributes</span></span>

<span data-ttu-id="441fc-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="441fc-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="441fc-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="441fc-106">Child elements</span></span>

|  <span data-ttu-id="441fc-107">Élément</span><span class="sxs-lookup"><span data-stu-id="441fc-107">Element</span></span>  |  <span data-ttu-id="441fc-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="441fc-108">Required</span></span>  |  <span data-ttu-id="441fc-109">Description</span><span class="sxs-lookup"><span data-stu-id="441fc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="441fc-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="441fc-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="441fc-111">Oui</span><span class="sxs-lookup"><span data-stu-id="441fc-111">Yes</span></span>  | <span data-ttu-id="441fc-112">Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="441fc-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="441fc-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="441fc-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
