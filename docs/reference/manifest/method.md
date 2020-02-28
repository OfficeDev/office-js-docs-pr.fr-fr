---
title: Élément Method dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2bcc24abf269f5d6c44c03e738bac480fd05d5ca
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324847"
---
# <a name="method-element"></a><span data-ttu-id="ce01c-102">Method, élément</span><span class="sxs-lookup"><span data-stu-id="ce01c-102">Method element</span></span>

<span data-ttu-id="ce01c-103">Spécifie une méthode individuelle de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="ce01c-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="ce01c-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="ce01c-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ce01c-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ce01c-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="ce01c-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ce01c-106">Contained in</span></span>

[<span data-ttu-id="ce01c-107">Méthodes</span><span class="sxs-lookup"><span data-stu-id="ce01c-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="ce01c-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="ce01c-108">Attributes</span></span>

|<span data-ttu-id="ce01c-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="ce01c-109">**Attribute**</span></span>|<span data-ttu-id="ce01c-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="ce01c-110">**Type**</span></span>|<span data-ttu-id="ce01c-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="ce01c-111">**Required**</span></span>|<span data-ttu-id="ce01c-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="ce01c-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ce01c-113">Nom</span><span class="sxs-lookup"><span data-stu-id="ce01c-113">Name</span></span>|<span data-ttu-id="ce01c-114">string</span><span class="sxs-lookup"><span data-stu-id="ce01c-114">string</span></span>|<span data-ttu-id="ce01c-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ce01c-115">required</span></span>|<span data-ttu-id="ce01c-116">Spécifie le nom de la méthode qualifiée requise avec son objet parent.</span><span class="sxs-lookup"><span data-stu-id="ce01c-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="ce01c-117">Par exemple, pour spécifier la `getSelectedDataAsync` méthode, vous devez spécifier `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="ce01c-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="ce01c-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="ce01c-118">Remarks</span></span>

<span data-ttu-id="ce01c-119">Les `Methods` éléments `Method` et ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de conditions requises, voir [versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce01c-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ce01c-120">Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément.</span><span class="sxs-lookup"><span data-stu-id="ce01c-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="ce01c-121">Pour plus d’informations sur la façon de procéder, consultez [la rubrique Understanding the Office JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="ce01c-121">For more information about how to do this, see [Understanding the Office JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

