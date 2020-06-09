---
title: Élément Script dans le fichier manifeste
description: L’élément script définit les paramètres de script qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 791f49f15673a029b982e40946f8cc90f02ba887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608089"
---
# <a name="script-element"></a><span data-ttu-id="c510c-103">Élément Script</span><span class="sxs-lookup"><span data-stu-id="c510c-103">Script element</span></span>

<span data-ttu-id="c510c-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="c510c-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c510c-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="c510c-105">Attributes</span></span>

<span data-ttu-id="c510c-106">Aucun</span><span class="sxs-lookup"><span data-stu-id="c510c-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="c510c-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c510c-107">Child elements</span></span>

|<span data-ttu-id="c510c-108">Éléments</span><span class="sxs-lookup"><span data-stu-id="c510c-108">Elements</span></span>  |  <span data-ttu-id="c510c-109">Requis</span><span class="sxs-lookup"><span data-stu-id="c510c-109">Required</span></span>  |  <span data-ttu-id="c510c-110">Description</span><span class="sxs-lookup"><span data-stu-id="c510c-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c510c-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c510c-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="c510c-112">Oui</span><span class="sxs-lookup"><span data-stu-id="c510c-112">Yes</span></span>  | <span data-ttu-id="c510c-113">Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c510c-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="c510c-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="c510c-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
