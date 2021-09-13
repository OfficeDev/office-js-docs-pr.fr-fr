---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de add-in est désactivée au lancement du module.
ms.date: 01/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a14385f7114eb3d35845b5d9873bdd718b46c0e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153051"
---
# <a name="enabled-element"></a>Élément Enabled

Spécifie si un [contrôle Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) est activé au lancement du module. **L’élément Enabled** est un élément enfant de [Control](control.md). S’il est omis, la valeur par défaut est `true` .

Cet élément n’est valide que dans Excel ; autrement dit, lorsque `Name` l’attribut de l’élément [Host](host.md) est « Workbook ».

Le contrôle parent peut également être activé et désactivé par programmation. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
