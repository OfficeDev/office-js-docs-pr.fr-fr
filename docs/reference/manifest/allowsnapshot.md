---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720959"
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
 > **AllowSnapshot** est défini sur `true` par défaut. Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.

