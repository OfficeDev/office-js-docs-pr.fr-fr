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
# <a name="allowsnapshot-element"></a>AllowSnapshot, élément

Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.

**Type de complément :** Contenu

## <a name="syntax"></a>Syntaxe

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

 > [!IMPORTANT]
 > **AllowSnapshot** est défini sur `true` par défaut. Cela rend une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application Office qui ne prend pas en charge les compléments Office, ou fournit une image statique du complément si l’application ne peut pas se connecter au serveur qui héberge le complément. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.
