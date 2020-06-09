---
title: Élément Page dans le fichier manifeste
description: L’élément page définit les paramètres de page HTML qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611497"
---
# <a name="page-element"></a><span data-ttu-id="84184-103">Élément Page</span><span class="sxs-lookup"><span data-stu-id="84184-103">Page element</span></span>

<span data-ttu-id="84184-104">Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="84184-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="84184-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="84184-105">Attributes</span></span>

<span data-ttu-id="84184-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="84184-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="84184-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="84184-107">Child elements</span></span>

|  <span data-ttu-id="84184-108">Élément</span><span class="sxs-lookup"><span data-stu-id="84184-108">Element</span></span>  |  <span data-ttu-id="84184-109">Requis</span><span class="sxs-lookup"><span data-stu-id="84184-109">Required</span></span>  |  <span data-ttu-id="84184-110">Description</span><span class="sxs-lookup"><span data-stu-id="84184-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="84184-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="84184-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="84184-112">Oui</span><span class="sxs-lookup"><span data-stu-id="84184-112">Yes</span></span>  | <span data-ttu-id="84184-113">Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="84184-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="84184-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="84184-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
