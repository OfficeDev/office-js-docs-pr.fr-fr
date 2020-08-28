---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294275"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="9ef00-103">AllowSnapshot, élément</span><span class="sxs-lookup"><span data-stu-id="9ef00-103">AllowSnapshot element</span></span>

<span data-ttu-id="9ef00-104">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="9ef00-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="9ef00-105">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="9ef00-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="9ef00-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9ef00-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="9ef00-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="9ef00-107">Contained in</span></span>

[<span data-ttu-id="9ef00-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9ef00-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="9ef00-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="9ef00-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="9ef00-110">**AllowSnapshot** est défini sur `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="9ef00-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="9ef00-111">Cela rend une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application Office qui ne prend pas en charge les compléments Office, ou fournit une image statique du complément si l’application ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="9ef00-111">This makes an image of the add-in visible for users that open the document in a version of the Office application that doesn't support Office Add-ins, or provides a static image of the add-in if the application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="9ef00-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="9ef00-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>
