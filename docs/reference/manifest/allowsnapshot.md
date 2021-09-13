---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 723817557020f4ec3dbe5b3135877fe49bf67acb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152188"
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
