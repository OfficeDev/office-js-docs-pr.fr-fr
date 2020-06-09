---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c46dcd882592c0b015dae4b9774533b96fe75cfe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608788"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="df9bf-103">AllowSnapshot, élément</span><span class="sxs-lookup"><span data-stu-id="df9bf-103">AllowSnapshot element</span></span>

<span data-ttu-id="df9bf-104">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="df9bf-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="df9bf-105">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="df9bf-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="df9bf-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="df9bf-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="df9bf-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="df9bf-107">Contained in</span></span>

[<span data-ttu-id="df9bf-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="df9bf-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="df9bf-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="df9bf-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="df9bf-110">**AllowSnapshot** est défini sur `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="df9bf-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="df9bf-111">Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="df9bf-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="df9bf-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="df9bf-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

