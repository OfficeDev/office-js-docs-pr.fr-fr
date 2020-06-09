---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611567"
---
# <a name="enabled-element"></a>Élément Enabled

Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément. L’élément **Enabled** est un élément enfant de [Control](control.md). Si ce paramètre est omis, la valeur par défaut est `true` .

Le contrôle parent peut également être activé et désactivé par programme. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
