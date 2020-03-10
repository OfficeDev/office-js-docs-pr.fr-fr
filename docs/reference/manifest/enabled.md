---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566202"
---
# <a name="enabled-element"></a>Élément Enabled

Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément. L’élément **Enabled** est un élément enfant de [Control](control.md). Si ce paramètre est omis, la valeur par `true`défaut est. 

Le contrôle parent peut également être activé et désactivé par programme. Pour plus d’informations, consultez la rubrique [activer et désactiver les commandes de complément](/office/dev/add-ins/design/disable-add-in-commands).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```

