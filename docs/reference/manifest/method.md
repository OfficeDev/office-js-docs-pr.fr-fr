---
title: Élément Method dans le fichier manifeste
description: L’élément Method spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de vos compléments Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c3531475a920fd24ce8390170b5f4728d4dcd0e0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611756"
---
# <a name="method-element"></a><span data-ttu-id="555be-103">Élément Method</span><span class="sxs-lookup"><span data-stu-id="555be-103">Method element</span></span>

<span data-ttu-id="555be-104">Spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="555be-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="555be-105">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="555be-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="555be-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="555be-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="555be-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="555be-107">Contained in</span></span>

[<span data-ttu-id="555be-108">Méthodes</span><span class="sxs-lookup"><span data-stu-id="555be-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="555be-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="555be-109">Attributes</span></span>

|<span data-ttu-id="555be-110">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="555be-110">**Attribute**</span></span>|<span data-ttu-id="555be-111">**Type**</span><span class="sxs-lookup"><span data-stu-id="555be-111">**Type**</span></span>|<span data-ttu-id="555be-112">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="555be-112">**Required**</span></span>|<span data-ttu-id="555be-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="555be-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="555be-114">Nom</span><span class="sxs-lookup"><span data-stu-id="555be-114">Name</span></span>|<span data-ttu-id="555be-115">string</span><span class="sxs-lookup"><span data-stu-id="555be-115">string</span></span>|<span data-ttu-id="555be-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="555be-116">required</span></span>|<span data-ttu-id="555be-117">Spécifie le nom de la méthode qualifiée requise avec son objet parent.</span><span class="sxs-lookup"><span data-stu-id="555be-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="555be-118">Par exemple, pour spécifier la `getSelectedDataAsync` méthode, vous devez spécifier `"Document.getSelectedDataAsync"` .</span><span class="sxs-lookup"><span data-stu-id="555be-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="555be-119">Remarques</span><span class="sxs-lookup"><span data-stu-id="555be-119">Remarks</span></span>

<span data-ttu-id="555be-120">Les `Methods` `Method` éléments et ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de conditions requises, voir [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="555be-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="555be-121">Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément.</span><span class="sxs-lookup"><span data-stu-id="555be-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="555be-122">Pour plus d’informations sur la façon de procéder, consultez [la rubrique Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="555be-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
