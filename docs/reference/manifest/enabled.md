---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de add-in est désactivée au lancement du module.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 54d28839a274ff41bab0b1e2cdd2d169e76c5815095950dec67ce2564eade601
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093904"
---
# <a name="enabled-element"></a>Élément Enabled

Spécifie si un [contrôle Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) est activé au lancement du module. **L’élément Enabled** est un élément enfant de [Control](control.md). S’il est omis, la valeur par défaut est `true` .

Cet élément n’est valide que dans Excel ; autrement dit, lorsque `Name` l’attribut de l’élément [Host](host.md) est « Workbook ».

Le contrôle parent peut également être activé et désactivé par programmation. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
