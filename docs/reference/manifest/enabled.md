---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de add-in est désactivée au lancement du module.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936880"
---
# <a name="enabled-element"></a>Élément Enabled

Spécifie si un [contrôle Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) est activé au lancement du module. **L’élément Enabled** est un élément enfant de [Control](control.md). S’il est omis, la valeur par défaut est `true` .

Cet élément n’est valide que dans Excel ; autrement dit, lorsque `Name` l’attribut de l’élément [Host](host.md) est « Workbook ».

Le contrôle parent peut également être activé et désactivé par programmation. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
