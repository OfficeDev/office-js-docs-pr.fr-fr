---
title: Élément Page dans le fichier manifeste
description: L’élément page définit les paramètres de page HTML qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720483"
---
# <a name="page-element"></a><span data-ttu-id="a6a1b-103">Élément Page</span><span class="sxs-lookup"><span data-stu-id="a6a1b-103">Page element</span></span>

<span data-ttu-id="a6a1b-104">Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a6a1b-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a6a1b-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="a6a1b-105">Attributes</span></span>

<span data-ttu-id="a6a1b-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="a6a1b-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a6a1b-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="a6a1b-107">Child elements</span></span>

|  <span data-ttu-id="a6a1b-108">Élément</span><span class="sxs-lookup"><span data-stu-id="a6a1b-108">Element</span></span>  |  <span data-ttu-id="a6a1b-109">Requis</span><span class="sxs-lookup"><span data-stu-id="a6a1b-109">Required</span></span>  |  <span data-ttu-id="a6a1b-110">Description</span><span class="sxs-lookup"><span data-stu-id="a6a1b-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a6a1b-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a6a1b-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a6a1b-112">Oui</span><span class="sxs-lookup"><span data-stu-id="a6a1b-112">Yes</span></span>  | <span data-ttu-id="a6a1b-113">Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a6a1b-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a6a1b-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6a1b-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
