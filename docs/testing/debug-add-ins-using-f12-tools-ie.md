---
title: Déboguer des compléments à l’aide des outils de développement pour Internet Explorer
description: Déboguer des applications à l’aide des outils de développement dans Internet Explorer.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb7c328e6b1f839d5d711f81beceaf7519545fe3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744670"
---
# <a name="debug-add-ins-using-developer-tools-in-internet-explorer"></a>Déboguer des applications à l’aide des outils de développement dans Internet Explorer

Cet article montre comment déboguer le code côté client (JavaScript ou TypeScript) de votre add-in lorsque les conditions suivantes sont remplies.

- Vous ne pouvez pas ou ne souhaitez pas déboguer à l’aide des outils intégrés à votre IDE ; ou vous rencontrez un problème qui se produit uniquement lorsque le module est exécuté en dehors de l’IDE.
- Votre ordinateur utilise une combinaison de versions Windows et Office qui utilisent le contrôle WebView Internet Explorer Trident.

Pour déterminer quel navigateur est utilisé sur votre ordinateur, voir [Navigateurs utilisés par les Office les autres.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Pour installer une version de Office qui utilise le webview Internet Explorer ou pour forcer votre version actuelle à utiliser Internet Explorer, voir Basculer vers le [webview Internet Explorer 11](#switch-to-the-internet-explorer-11-webview).

## <a name="debug-a-task-pane-add-in-using-the-f12-tools"></a>Déboguer un add-in du volet Des tâches à l’aide des outils F12

Windows 10 11 incluent un outil de développement web appelé « F12 », car il a été lancé à l’origine en appuyant sur F12 dans Internet Explorer. F12 est désormais une application indépendante utilisée pour déboguer votre application lorsqu’elle est en cours d’exécution dans le contrôle WebView Internet Explorer Trident. L’application n’est pas disponible dans les versions antérieures de Windows.

> [!NOTE]
> Si votre add-in dispose d’une commande de add-in qui exécute une fonction, [celle-ci](../design/add-in-commands.md) s’exécute dans un processus de navigateur masqué que les outils F12 ne peuvent pas détecter ni attacher. La technique décrite dans cet article ne peut donc pas être utilisée pour déboguer du code dans la fonction.

Les étapes suivantes sont les instructions pour le débogage de votre add-in. Si vous souhaitez simplement tester les outils F12 eux-mêmes, voir exemple de [add-in pour tester les outils F12](#example-add-in-to-test-the-f12-tools).

1. [Chargez une](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) version de version et exécutez le module.
1. Lancez les outils de développement F12 qui correspondent à votre version de Office.

   - Pour la version 32 bits, utilisez C:\Windows\System32\F12\IEChooser.exe
   - Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\IEChooser.exe

   IEChooser s’ouvre avec une fenêtre nommée **Choisir la cible à déboguer**. Votre add-in apparaît dans la fenêtre nommée par le nom de fichier de la page d’accueil du module. Dans la capture d’écran suivante, il s’agit de `Home.html`. Seuls les processus en cours d’exécution dans Internet Explorer, ou Trident, apparaissent. L’outil ne peut pas être attaché aux processus en cours d’exécution dans d’autres navigateurs ou vues web, y compris Microsoft Edge.

    :::image type="content" source="../images/choose-target-to-debug.png" alt-text="Écran IEChooser, avec plusieurs processus Internet Explorer et Trident répertoriés. L’un d’entre Home.html.":::

1. Sélectionnez le processus de votre add-in . autrement dit, son nom de fichier de page d’accueil. Cette action attache les outils F12 au processus et ouvre l’interface utilisateur F12 principale.
1. Ouvrez l’onglet **Débogueur**.
1. Dans le coin supérieur gauche de l’onglet, juste en dessous du ruban de l’outil débogger, se trouve une petite icône de dossier. Sélectionnez cette valeur pour ouvrir une liste de listes listes des fichiers dans le module. Voici un exemple.

    :::image type="content" source="../images/f12-file-dropdown.png" alt-text="Capture d’écran du coin supérieur gauche de l’onglet débogger avec une liste de dossiers ouverte et une liste de fichiers.":::

1. Sélectionnez le fichier à déboguer et celui-ci s’ouvre dans le **volet de script** (gauche) de l’onglet **Déboguer** . Si vous utilisez un transpiler, un bundler ou un minifier qui modifie le nom du fichier, il aura le nom final qui est effectivement chargé, et non le nom du fichier source d’origine.

1. Faites défiler jusqu’à une ligne où vous souhaitez définir un point d’arrêt et cliquez dans la marge à gauche du numéro de ligne. Vous verrez un point rouge à gauche de la ligne et une ligne correspondante apparaît dans l’onglet Points  d’arrêt du volet inférieur droit. La capture d'écran suivante présente un exemple :

    :::image type="content" source="../images/debugger-home-js-02.png" alt-text="Débogger avec point d’arrêt home.js fichier.":::

1. Exécutez les fonctions dans le complément, si nécessaire, afin de déclencher le point d’arrêt. Lorsque le point d’arrêt est atteint, une flèche pointant vers la droite apparaît sur le point rouge du point d’arrêt. La capture d'écran suivante présente un exemple :

    :::image type="content" source="../images/debugger-home-js-01.png" alt-text="Débogger avec les résultats du point d’arrêt déclenché.":::

> [!TIP]
> Pour plus d’informations sur l’utilisation des outils F12, voir Inspecter en javaScript en cours d’exécution [avec le débogger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).

### <a name="example-add-in-to-test-the-f12-tools"></a>Exemple de add-in pour tester les outils F12

Cet exemple utilise Word et un complément gratuit d’AppSource.

1. Ouvrez un document vierge dans Word. 
1. Sous **l’onglet** Insertion, dans le groupe Des **modules**, sélectionnez Mes applications pour ouvrir la boîte de  dialogue Office, puis sélectionnez l’onglet **STORE**.
1. Sélectionnez **le add-in QR4Office** . Il s’ouvre dans un volet Des tâches.
1. Lancez les outils de développement F12 qui correspondent à votre version de Office comme décrit dans la section précédente.
1. Dans la fenêtre F12, **sélectionnezHome.html**.
1. Dans **l’onglet Débogger** , ouvrez le **fichierHome.js** comme décrit dans la section précédente.
1. Définissez les points d’arrêt sur les lignes 310 et 312.
1. Dans le module, sélectionnez le **bouton Insérer** . L’un ou l’autre point d’arrêt est atteint.

## <a name="debug-a-dialog-in-an-add-in"></a>Débogage d’une boîte de dialogue dans un add-in

Si votre application utilise l’API de boîte de dialogue Office, la boîte de dialogue s’exécute dans un processus distinct du volet Des tâches (le cas cas), et les outils doivent s’attacher à ce processus. Procédez comme suit.

1. Exécutez le module et les outils. 
1. Ouvrez la boîte de dialogue, puis sélectionnez **le bouton Actualiser** dans les outils. Le processus de boîte de dialogue s’affiche. Son nom est le nom du fichier qui est ouvert dans la boîte de dialogue.
1. Sélectionnez le processus pour l’ouvrir et déboguer, comme décrit dans la section Déboguer un add-in du volet Des tâches à l’aide des outils [F12](#debug-a-task-pane-add-in-using-the-f12-tools).

## <a name="switch-to-the-internet-explorer-11-webview"></a>Basculer vers le webview Internet Explorer 11

Il existe deux façons de changer le mode web d’Internet Explorer. Vous pouvez exécuter une commande simple dans une invite de commandes ou installer une version de Office qui utilise Internet Explorer par défaut. Nous vous recommandons la première méthode. Mais vous devez utiliser le deuxième scénario dans les scénarios suivants.

- Votre projet a été développé avec Visual Studio et IIS. Il n’est pas node.js base.
- Vous souhaitez être absolument robuste dans vos tests.
- Si, pour une raison quelconque, l’outil de ligne de commande ne fonctionne pas.

### <a name="switch-via-the-command-line"></a>Basculer via la ligne de commande

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Installer une version de Office qui utilise Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Voir aussi

- [Inspecter le code JavaScript en cours d’exécution avec le débogueur](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Utilisation des outils de développement F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
