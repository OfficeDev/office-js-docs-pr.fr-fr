---
title: Élément Page dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433732"
---
# <a name="page-element"></a><span data-ttu-id="5974c-102">Élément Page</span><span class="sxs-lookup"><span data-stu-id="5974c-102">Page element</span></span>

<span data-ttu-id="5974c-103">Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5974c-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5974c-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="5974c-104">Attributes</span></span>

<span data-ttu-id="5974c-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="5974c-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5974c-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="5974c-106">Child elements</span></span>

|  <span data-ttu-id="5974c-107">Élément</span><span class="sxs-lookup"><span data-stu-id="5974c-107">Element</span></span>  |  <span data-ttu-id="5974c-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5974c-108">Required</span></span>  |  <span data-ttu-id="5974c-109">Description</span><span class="sxs-lookup"><span data-stu-id="5974c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5974c-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5974c-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5974c-111">Oui</span><span class="sxs-lookup"><span data-stu-id="5974c-111">Yes</span></span>  | <span data-ttu-id="5974c-112">Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5974c-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="5974c-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="5974c-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
