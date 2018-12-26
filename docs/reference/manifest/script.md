---
title: Élément Script dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433137"
---
# <a name="script-element"></a><span data-ttu-id="5dbeb-102">Élément Script</span><span class="sxs-lookup"><span data-stu-id="5dbeb-102">Script element</span></span>

<span data-ttu-id="5dbeb-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5dbeb-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5dbeb-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="5dbeb-104">Attributes</span></span>

<span data-ttu-id="5dbeb-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="5dbeb-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5dbeb-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="5dbeb-106">Child elements</span></span>

|<span data-ttu-id="5dbeb-107">Éléments</span><span class="sxs-lookup"><span data-stu-id="5dbeb-107">Elements</span></span>  |  <span data-ttu-id="5dbeb-108">Requis</span><span class="sxs-lookup"><span data-stu-id="5dbeb-108">Required</span></span>  |  <span data-ttu-id="5dbeb-109">Description</span><span class="sxs-lookup"><span data-stu-id="5dbeb-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5dbeb-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5dbeb-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5dbeb-111">Oui</span><span class="sxs-lookup"><span data-stu-id="5dbeb-111">Yes</span></span>  | <span data-ttu-id="5dbeb-112">Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5dbeb-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="5dbeb-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="5dbeb-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
