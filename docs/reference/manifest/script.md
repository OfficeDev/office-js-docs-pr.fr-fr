---
title: Élément Script dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8352ada0eeb6af071d5f20f750dcdeaefe31e918
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450435"
---
# <a name="script-element"></a><span data-ttu-id="f37c7-102">Élément Script</span><span class="sxs-lookup"><span data-stu-id="f37c7-102">Script element</span></span>

<span data-ttu-id="f37c7-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="f37c7-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f37c7-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="f37c7-104">Attributes</span></span>

<span data-ttu-id="f37c7-105">Aucun</span><span class="sxs-lookup"><span data-stu-id="f37c7-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="f37c7-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="f37c7-106">Child elements</span></span>

|<span data-ttu-id="f37c7-107">Éléments</span><span class="sxs-lookup"><span data-stu-id="f37c7-107">Elements</span></span>  |  <span data-ttu-id="f37c7-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="f37c7-108">Required</span></span>  |  <span data-ttu-id="f37c7-109">Description</span><span class="sxs-lookup"><span data-stu-id="f37c7-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f37c7-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f37c7-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="f37c7-111">Oui</span><span class="sxs-lookup"><span data-stu-id="f37c7-111">Yes</span></span>  | <span data-ttu-id="f37c7-112">Chaîne avec l’ID de ressource du fichier JavaScript utilisé par les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f37c7-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="f37c7-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="f37c7-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
