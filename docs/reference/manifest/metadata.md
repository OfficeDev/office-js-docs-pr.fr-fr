---
title: Élément de métadonnées dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452045"
---
# <a name="metadata-element"></a><span data-ttu-id="b8b5b-102">Élément de métadonnées</span><span class="sxs-lookup"><span data-stu-id="b8b5b-102">Metadata element</span></span>

<span data-ttu-id="b8b5b-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b8b5b-103">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="b8b5b-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="b8b5b-104">Attributes</span></span>

<span data-ttu-id="b8b5b-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="b8b5b-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="b8b5b-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b8b5b-106">Child elements</span></span>

|  <span data-ttu-id="b8b5b-107">Élément</span><span class="sxs-lookup"><span data-stu-id="b8b5b-107">Element</span></span>  |  <span data-ttu-id="b8b5b-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="b8b5b-108">Required</span></span>  |  <span data-ttu-id="b8b5b-109">Description</span><span class="sxs-lookup"><span data-stu-id="b8b5b-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b8b5b-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b8b5b-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="b8b5b-111">Oui</span><span class="sxs-lookup"><span data-stu-id="b8b5b-111">Yes</span></span>  | <span data-ttu-id="b8b5b-112">Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b8b5b-112">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="b8b5b-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="b8b5b-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
