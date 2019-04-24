---
title: Élément Page dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452073"
---
# <a name="page-element"></a><span data-ttu-id="3ef36-102">Élément Page</span><span class="sxs-lookup"><span data-stu-id="3ef36-102">Page element</span></span>

<span data-ttu-id="3ef36-103">Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="3ef36-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="3ef36-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="3ef36-104">Attributes</span></span>

<span data-ttu-id="3ef36-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="3ef36-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="3ef36-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="3ef36-106">Child elements</span></span>

|  <span data-ttu-id="3ef36-107">Élément</span><span class="sxs-lookup"><span data-stu-id="3ef36-107">Element</span></span>  |  <span data-ttu-id="3ef36-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3ef36-108">Required</span></span>  |  <span data-ttu-id="3ef36-109">Description</span><span class="sxs-lookup"><span data-stu-id="3ef36-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3ef36-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3ef36-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="3ef36-111">Oui</span><span class="sxs-lookup"><span data-stu-id="3ef36-111">Yes</span></span>  | <span data-ttu-id="3ef36-112">Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="3ef36-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="3ef36-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="3ef36-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
