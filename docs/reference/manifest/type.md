---
title: Élément type dans le fichier manifeste
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356873"
---
# <a name="type-element"></a><span data-ttu-id="085c7-102">Élément Type</span><span class="sxs-lookup"><span data-stu-id="085c7-102">Type element</span></span>

<span data-ttu-id="085c7-103">Indique si le complément équivalent est un complément COM ou un XLL.</span><span class="sxs-lookup"><span data-stu-id="085c7-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="085c7-104">**Type de complément:** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="085c7-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="085c7-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="085c7-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="085c7-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="085c7-106">Contained in</span></span>

[<span data-ttu-id="085c7-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="085c7-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="085c7-108">Valeurs de type de complément</span><span class="sxs-lookup"><span data-stu-id="085c7-108">Add-in type values</span></span>

<span data-ttu-id="085c7-109">Vous devez spécifier l'une des valeurs suivantes pour l' `Type` élément.</span><span class="sxs-lookup"><span data-stu-id="085c7-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="085c7-110">COM: spécifie que le complément équivalent est un complément COM.</span><span class="sxs-lookup"><span data-stu-id="085c7-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="085c7-111">XLL: spécifie que le complément équivalent est une XLL Excel.</span><span class="sxs-lookup"><span data-stu-id="085c7-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="085c7-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="085c7-112">See also</span></span>

- [<span data-ttu-id="085c7-113">Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur</span><span class="sxs-lookup"><span data-stu-id="085c7-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="085c7-114">Faire en sorte que votre complément Office soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="085c7-114">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)