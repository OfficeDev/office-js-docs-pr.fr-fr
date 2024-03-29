---
title: Attacher un débogueur à partir du volet Office
description: Découvrez comment attacher un débogger à partir du volet Des tâches.
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0363b7966ab3da11167cb4b0cd324df28fd9efb3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744749"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Attacher un débogueur à partir du volet Office

Dans certains environnements, un débogger peut être attaché à un Office qui est déjà en cours d’exécution. Cela peut être utile lorsque vous souhaitez déboguer un add-in qui est déjà en transit ou en production. Si vous développez et testez encore le add-in, voir Vue d’ensemble du débogage [Office des modules.](debug-add-ins-overview.md)

La technique décrite dans cet article ne peut être utilisée que lorsque les conditions suivantes sont remplies.

- Le module est en cours d’exécution Office sur Windows.
- L’ordinateur utilise une combinaison de versions Windows et Office qui utilisent le contrôle WebView Edge (basé sur Chromium), WebView2. Pour déterminer le navigateur que vous utilisez, consultez Les [navigateurs utilisés par les Office des modules.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Pour lancer le débompeur, choisissez le coin supérieur droit du volet Des tâches pour activer **le menu Personnalité** (comme illustré dans le cercle rouge de l’image suivante).

![Capture d’écran du menu Attacher le débogger.](../images/attach-debugger.png)

Sélectionnez **Attacher le débogueur**. Cela lance les outils Microsoft Edge de développement Chromium (basés sur un logiciel). Utilisez les outils comme décrit dans Déboguer des applications à l’aide des [outils de développement Microsoft Edge (Chromium)](debug-add-ins-using-devtools-edge-chromium.md).

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble du débogage Office des modules](debug-add-ins-overview.md)
