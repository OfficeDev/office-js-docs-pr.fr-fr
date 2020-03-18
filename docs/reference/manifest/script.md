---
title: Élément Script dans le fichier manifeste
description: L’élément script définit les paramètres de script qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f05fc85bd0454c340f4352bb73f299b9e7730224
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720413"
---
# <a name="script-element"></a><span data-ttu-id="7e9b9-103">Élément Script</span><span class="sxs-lookup"><span data-stu-id="7e9b9-103">Script element</span></span>

<span data-ttu-id="7e9b9-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="7e9b9-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="7e9b9-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="7e9b9-105">Attributes</span></span>

<span data-ttu-id="7e9b9-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="7e9b9-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="7e9b9-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7e9b9-107">Child elements</span></span>

|<span data-ttu-id="7e9b9-108">Éléments</span><span class="sxs-lookup"><span data-stu-id="7e9b9-108">Elements</span></span>  |  <span data-ttu-id="7e9b9-109">Requis</span><span class="sxs-lookup"><span data-stu-id="7e9b9-109">Required</span></span>  |  <span data-ttu-id="7e9b9-110">Description</span><span class="sxs-lookup"><span data-stu-id="7e9b9-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7e9b9-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7e9b9-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="7e9b9-112">Oui</span><span class="sxs-lookup"><span data-stu-id="7e9b9-112">Yes</span></span>  | <span data-ttu-id="7e9b9-113">Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="7e9b9-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="7e9b9-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="7e9b9-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
