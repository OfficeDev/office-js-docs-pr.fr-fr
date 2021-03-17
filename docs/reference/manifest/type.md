---
title: Élément Type dans le fichier manifeste
description: L’élément Type spécifie si le compl?ment équivalent est un compl?ment COM ou un XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836808"
---
# <a name="type-element"></a><span data-ttu-id="8154e-103">Élément Type</span><span class="sxs-lookup"><span data-stu-id="8154e-103">Type element</span></span>

<span data-ttu-id="8154e-104">Spécifie si le compl?ment équivalent est un compl?ment COM ou un XLL.</span><span class="sxs-lookup"><span data-stu-id="8154e-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="8154e-105">**Type de add-in :** Volet Des tâches, Fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="8154e-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="8154e-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8154e-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="8154e-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8154e-107">Contained in</span></span>

[<span data-ttu-id="8154e-108">EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="8154e-108">EquivalentAddin</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="8154e-109">Valeurs de type de add-in</span><span class="sxs-lookup"><span data-stu-id="8154e-109">Add-in type values</span></span>

<span data-ttu-id="8154e-110">Vous devez spécifier l’une des valeurs suivantes pour `Type` l’élément.</span><span class="sxs-lookup"><span data-stu-id="8154e-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="8154e-111">COM : spécifie que le compl?ment équivalent est un compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="8154e-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="8154e-112">XLL : spécifie que le add-in équivalent est une XLL Excel.</span><span class="sxs-lookup"><span data-stu-id="8154e-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="8154e-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8154e-113">See also</span></span>

- [<span data-ttu-id="8154e-114">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="8154e-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="8154e-115">Rendre votre complément Office compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="8154e-115">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)