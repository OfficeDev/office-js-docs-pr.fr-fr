---
title: Déboguer des compléments Office dans Visual Studio
description: Utiliser Visual Studio pour déboguer des compléments Office dans le client de bureau Office sous Windows
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 7c49e3019c22af0b5d44a382b33187e5d2de4ceb
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430477"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Déboguer des compléments Office dans Visual Studio

Cet article décrit comment utiliser Visual Studio 2019 pour déboguer un complément Office dans le client de bureau Office sur Windows. Si vous utilisez une autre version de Visual Studio, les procédures peuvent légèrement varier. 

> [!NOTE]
> Vous ne pouvez pas utiliser Visual Studio pour déboguer des compléments Office dans Office sur le web ou sur Mac. Pour plus d’informations sur le débogage sur ces plateformes, voir [Déboguer les compléments Office dans Office sur le web](../testing/debug-add-ins-in-office-online.md) ou [Déboguer les compléments Office sur Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

## <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a>Activer le débogage pour les commandes de compléments et les codes sans interface utilisateur

Lors du débogage d’Office sous Windows par Visual Studio, le complément est hébergé dans une instance du navigateur Microsoft Internet Explorer ou Microsoft Edge. Pour identifier le navigateur utilisé sur votre ordinateur de développement, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).
> [!NOTE]
> La variable d'environnement JS_Debug n'est plus nécessaire dans la procédure ci-après. Pour plus d’informations, voir [Comportements de débogage dans les compléments web Office](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html) sur le forum de support de la Communauté des développeurs Microsoft.

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

## <a name="review-the-build-and-debug-properties"></a>Réviser les propriétés de création et de débogage

Avant de commencer le débogage, passez en revue les propriétés de chaque projet pour vérifier que Visual Studio ouvre l’application Office souhaitée et que les autres propriétés de génération et de débogage sont définies correctement.

### <a name="add-in-project-properties"></a>Propriétés du projet de complément

Ouvrir la fenêtre **Propriétés** pour le projet complément pour examiner les propriétés du projet :

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet de complément.

|**Propriété**|**Description**|
|:-----|:-----|
|**Action de démarrage**|Spécifie le mode de débogage pour votre complément. Actuellement, seul le mode **Client Office pour bureau** est pris en charge pour les projets complément Office.|
|**Document de démarrage**<br/>(Compléments Excel, PowerPoint et Word uniquement)|Spécifie le document à ouvrir lors du démarrage du projet.|
|**Projet Web**|Spécifie le nom du projet web associé au complément.|
|**Adresse e-mail**<br/>(Compléments Outlook uniquement)|Spécifie l’adresse de messagerie du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**Url EWS**<br/>(Compléments Outlook uniquement)|URL de service web Exchange (par exemple :`https://www.contoso.com/ews/exchange.aspx`). |
|**Url OWA**<br/>(Compléments Outlook uniquement)|URL Outlook sur le web (par exemple : `https://www.contoso.com/owa`).|
|**Utiliser l’authentification multi-facteur**<br/>(Compléments Outlook uniquement)|Valeur booléenne qui indique si l’authentification multi-facteur doit être utilisée.|
|**Nom d'utilisateur**<br/>(Compléments Outlook uniquement)|Spécifie le nom du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Emplacement du fichier de projet.|

> [!NOTE]
> Pour un complément Outlook, vous pouvez choisir de spécifier des valeurs pour une ou plusieurs des propriétés du*complément Outlook uniquement*dans la fenêtre**propriétés**, mais cette opération n’est pas obligatoire.

### <a name="web-application-project-properties"></a>Propriétés du projet application Web

Ouvrir la fenêtre**Propriétés** pour le projet complément web pour examiner les propriétés du projet :

1. Dans l' **Explorateur de solutions**, choisissez le projet d’application Web.

2. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet d’application web qui sont les plus pertinentes aux projets complément Office.

|**Propriété**|**Description**|
|:-----|:-----|
|**SSL activé**|Spécifie si SSL est activé sur le site. Cette propriété doit être définie sur **vrai** pour les projets complément Office.|
|**URL SSL**|Spécifie l’URL HTTPS sécurité pour le site. Lecture seule.|
|**URL**|Spécifie l’URL HTTP pour le site. Lecture seule.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Précise l’emplacement du fichier de projet. Lecture seule. Le fichier manifeste créé par Visual Studio lors de l’exécution est écrit le `bin\Debug\OfficeAppManifests` dossier dans cet emplacement.|

## <a name="use-an-existing-document-to-debug-the-add-in"></a>Utiliser un document existant pour déboguer le complément

Si vous avez un document qui contient les données de test à utiliser pendant le débogage de votre complément Excel, PowerPoint ou Word, Visual Studio peut être configuré pour ouvrir ce dernier lorsque vous démarrez le projet. Pour spécifier un document existant à utiliser pour déboguer le complément, procédez comme suit.

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menus, sélectionnez **Project** > **ajouter un élément existant**.

3. Dans la boîte de dialogue **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.

4. Choisissez le bouton**Ajouter** pour ajouter le document à votre projet.

5. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

6. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

7. Dans la fenêtre**propriétés**, choisissez la liste**Document de démarrage** et sélectionnez le document que vous avez ajouté au projet. Le projet est désormais configuré pour démarrer le complément dans ce document.

## <a name="start-the-project"></a>Démarrer le projet

Démarrez le projet en choisissant **déboguer** > **démarrer le débogage** à partir de la barre de menus. Visual Studio créera automatiquement la solution et démarrera Office pour héberger votre complément.

> [!NOTE]
> Lorsque vous commencez un projet de complément Outlook, vous serez invité à indiquer vos informations de connexion. Si vous êtes invité à vous connecter à plusieurs reprises ou si vous recevez un message d’erreur indiquant que vous n’êtes pas autorisé, l’authentification de base peut être désactivée pour les comptes sur votre client Microsoft 365. Dans ce cas, essayez d’utiliser un compte Microsoft à la place. Il se peut également que vous deviez définir la propriété « Utiliser l’authentification multifacteur » sur Vrai dans la boîte de dialogue Propriétés du complément Outlook Web.

Visual Studio génère le projet et effectue les actions suivantes :

1. Crée une copie du fichier manifeste XML et ajoute celui-ci au `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire. L’application Office qui héberge votre complément utilise cette copie lorsque vous démarrez Visual Studio et déboguez le complément.

2. Crée un ensemble d’entrées de Registre sur votre ordinateur qui permettent au complément d’apparaître dans l’application Office.

3. Génère le projet d’application web, puis le déploie sur le serveur web IIS local (https://localhost).

4. S’il s’agit du premier projet de complément que vous déployez sur un serveur web IIS local, il se peut que vous soyez invité à installer un certificat auto-signé pour le magasin de certificats racines de confiance de l’utilisateur actuel. Cela est nécessaire pour qu’IIS Express puisse afficher correctement le contenu de votre complément.

> [!NOTE]
> La dernière version d’Office peut utiliser un contrôle web plus récent pour afficher le contenu du complément lors de l’exécution de celui-ci sur Windows 10. Si tel est le cas, Visual Studio peut vous inviter à ajouter une exemption de bouclage de réseau local. Cela est nécessaire pour que le contrôle Web, dans l’application cliente Office, puisse accéder au site Web déployé sur le serveur Web IIS local. Vous pouvez également modifier ce paramètre à tout moment dans Visual Studio sous **Outils** > **Options** > **Outils Office (web)** > **Débogage de compléments web**.

Visual Studio effectue ensuite les actions suivantes :

1. Il modifie l’élément [SourceLocation](../reference/manifest/sourcelocation.md) du fichier manifeste XML en remplaçant le jeton`~remoteAppUrl`par l’adresse complète de la page de démarrage (par exemple`https://localhost:44302/Home.html`).

2. Il démarre le projet d’application web dans IIS Express.

3. Ouvre l’application Office.

Visual Studio n’affiche pas les erreurs de validation dans la fenêtre **OUTPUT** lorsque vous générez le projet. Visual Studio signale les erreurs et avertissements dans la fenêtre **ERRORLIST** lorsqu’elles se produisent. Visual Studio signale également des erreurs de validation en affichant les soulignements ondulés de différentes couleurs (également connus sous soulignements ondulés) dans l’éditeur de code et de texte. Ces marques signalent l’arrivée de problèmes Visual Studio détectés dans votre code. Pour plus d’informations sur comment activer ou désactiver la validation, voir [Options, éditeur de texte, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019&preserve-view=true).

Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

## <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Déboguer le code d’un complément Excel, PowerPoint ou Word

Si votre complément n’est pas visible dans le document qui est affiché dans l’application Office (Excel, PowerPoint ou Word) après [le démarrage du projet](#start-the-project), lancez manuellement le complément dans l’application Office. Par exemple, démarrez votre complément volet tâche en choisissant le bouton**Afficher le volet de tâches** dans l’onglet **Accueil**. Une fois que votre complément est affiché dans Excel, PowerPoint ou Word, vous pouvez déboguer votre code en procédant comme suit :

1. Dans Excel, PowerPoint ou Word, sélectionnez l’onglet **insérer**, puis cliquez sur la flèche vers le bas située à droite de **Mes compléments**.

    ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

2. Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez votre complément pour effectuer cette opération.

3. Dans Visual Studio, définissez des points d’arrêt dans votre code.

4. Dans Excel, PowerPoint ou Word, interagissez avec votre complément.

5. Lorsque des points d’arrêt sont marqués dans Visual Studio, parcourez le code si besoin.

Vous pouvez modifier votre code et passer en revue les effets de ces modifications dans votre complément sans avoir à fermer l’application Office et redémarrer le projet. Une fois que vous avez enregistré les modifications apportées à votre code, rechargez simplement le complément dans l’application Office. Par exemple, rechargez un complément de volet de tâches en choisissant le coin supérieur droit du volet Office pour activer la [menu personnalisé](../design/task-pane-add-ins.md#personality-menu), puis**Recharger**.

## <a name="debug-the-code-for-an-outlook-add-in"></a>Déboguer le code d’un complément Outlook

Une fois que vous avez [démarré le projet](#start-the-project) et que Visual Studio lance Outlook pour héberger votre complément, ouvrez un élément de courrier électronique ou un rendez-vous.

Outlook active le complément pour l’élément à condition que les critères d’activation soient respectés. La barre complément apparaît en haut de la fenêtre de l’inspecteur ou du volet de lecture, et votre complément Outlook apparaît sous la forme d’un bouton dans la barre du complément. Si votre complément est doté d’une commande, un bouton apparaît dans le ruban (soit dans l’onglet par défaut, soit dans un onglet personnalisé indiqué), et le complément n’apparaît pas dans la barre complément.

Pour voir votre complément Outlook, cliquez sur le bouton correspondant. Une fois que votre complément est affiché dans Outlook, vous pouvez déboguer votre code en procédant comme suit :

1. Dans Visual Studio, définissez des points d’arrêt dans votre code.

2. Dans Outlook, interagissez avec votre complément.

3. Lorsque des points d’arrêt sont marqués dans Visual Studio, parcourez le code si besoin.

Vous pouvez modifier votre code et passer en revue les effets de ces modifications dans votre complément sans avoir à fermer Outlook et redémarrer le projet. Une fois que vous enregistrez des modifications à votre code, il vous suffit d’ouvrir le menu contextuel pour le complément (dans Outlook), puis **recharger**.

## <a name="next-steps"></a>Étapes suivantes

Une fois que votre complément fonctionne comme vous le souhaitez, voir [Déployer et publier votre complément Office](../publish/publish.md) pour en savoir plus sur les méthodes avec lesquelles vous pouvez distribuer le complément aux utilisateurs.
