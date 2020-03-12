---
title: Élément Method dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 74b7a8b3d0f8511d21eb0df150500850e8b93fe9
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596892"
---
# <a name="method-element"></a><span data-ttu-id="82e54-102">Method, élément</span><span class="sxs-lookup"><span data-stu-id="82e54-102">Method element</span></span>

<span data-ttu-id="82e54-103">Spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="82e54-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="82e54-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="82e54-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="82e54-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="82e54-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="82e54-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="82e54-106">Contained in</span></span>

[<span data-ttu-id="82e54-107">Méthodes</span><span class="sxs-lookup"><span data-stu-id="82e54-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="82e54-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="82e54-108">Attributes</span></span>

|<span data-ttu-id="82e54-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="82e54-109">**Attribute**</span></span>|<span data-ttu-id="82e54-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="82e54-110">**Type**</span></span>|<span data-ttu-id="82e54-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="82e54-111">**Required**</span></span>|<span data-ttu-id="82e54-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="82e54-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="82e54-113">Nom</span><span class="sxs-lookup"><span data-stu-id="82e54-113">Name</span></span>|<span data-ttu-id="82e54-114">string</span><span class="sxs-lookup"><span data-stu-id="82e54-114">string</span></span>|<span data-ttu-id="82e54-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="82e54-115">required</span></span>|<span data-ttu-id="82e54-116">Spécifie le nom de la méthode qualifiée requise avec son objet parent.</span><span class="sxs-lookup"><span data-stu-id="82e54-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="82e54-117">Par exemple, pour spécifier la `getSelectedDataAsync` méthode, vous devez spécifier `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="82e54-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="82e54-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="82e54-118">Remarks</span></span>

<span data-ttu-id="82e54-119">Les `Methods` éléments `Method` et ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de conditions requises, voir [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="82e54-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="82e54-120">Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément.</span><span class="sxs-lookup"><span data-stu-id="82e54-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="82e54-121">Pour plus d’informations sur la façon de procéder, consultez [la rubrique Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="82e54-121">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
