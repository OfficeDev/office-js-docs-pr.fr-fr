---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771388"
---
# <a name="enabled-element"></a>Élément Enabled

Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément. L’élément **Enabled** est un élément enfant de [Control](control.md). Si ce paramètre est omis, la valeur par défaut est `true` .

Cet élément est valide uniquement dans Excel ; autrement dit, lorsque l' `Name` attribut de l’élément [hôte](host.md) est « classeur ».

Le contrôle parent peut également être activé et désactivé par programme. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
