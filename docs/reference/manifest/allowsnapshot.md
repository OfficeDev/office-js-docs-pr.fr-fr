---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937703"
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
 > **AllowSnapshot** est défini sur `true` par défaut. Cela rend une image du add-in visible pour les utilisateurs qui ouvrent le document dans une version de l’application Office qui ne prend pas en charge les Office Add-ins, ou fournit une image statique du module si l’application ne peut pas se connecter au serveur hébergeant le module. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.
