---
title: Élément Method dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 19234b35e1faf8a8cc52a9e893fcc720793cadae
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450652"
---
# <a name="method-element"></a><span data-ttu-id="19484-102">Method, élément</span><span class="sxs-lookup"><span data-stu-id="19484-102">Method element</span></span>

<span data-ttu-id="19484-103">Spécifie une méthode individuelle de l’API JavaScript pour Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="19484-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="19484-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="19484-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="19484-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="19484-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="19484-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="19484-106">Contained in</span></span>

[<span data-ttu-id="19484-107">Méthodes</span><span class="sxs-lookup"><span data-stu-id="19484-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="19484-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="19484-108">Attributes</span></span>

|<span data-ttu-id="19484-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="19484-109">**Attribute**</span></span>|<span data-ttu-id="19484-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="19484-110">**Type**</span></span>|<span data-ttu-id="19484-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="19484-111">**Required**</span></span>|<span data-ttu-id="19484-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="19484-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="19484-113">Nom</span><span class="sxs-lookup"><span data-stu-id="19484-113">Name</span></span>|<span data-ttu-id="19484-114">string</span><span class="sxs-lookup"><span data-stu-id="19484-114">string</span></span>|<span data-ttu-id="19484-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="19484-115">required</span></span>|<span data-ttu-id="19484-p101">Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="19484-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="19484-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="19484-118">Remarks</span></span>

<span data-ttu-id="19484-119">Les éléments **Methods** et **Method** ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de spécifications, voir l’article [Versions Office et jeux de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="19484-119">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="19484-120">Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément.</span><span class="sxs-lookup"><span data-stu-id="19484-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="19484-121">Pour plus d’informations sur la procédure à suivre, consultez l’article décrivant l’[API JavaScript pour Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="19484-121">For more information about how to do this, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

