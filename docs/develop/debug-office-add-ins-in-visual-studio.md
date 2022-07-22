---
title: Déboguer des compléments Office dans Visual Studio
description: Utilisez Visual Studio pour déboguer des compléments Office dans le client de bureau Office sur Windows.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 09693f81c069aba97740265fa88bf117a937c742
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958712"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Déboguer des compléments Office dans Visual Studio

Cet article explique comment déboguer du code côté client dans les compléments Office créés avec l’un des modèles de projet de complément Office dans Visual Studio 2022.  Pour plus d’informations sur le débogage du code côté serveur dans les compléments Office, consultez [Vue d’ensemble du débogage des compléments Office côté serveur ou côté client.](../testing/debug-add-ins-overview.md#server-side-or-client-side)

> [!NOTE]
> Vous ne pouvez pas utiliser Visual Studio pour déboguer des compléments dans Office sur Mac. Pour plus d’informations sur le débogage sur un Mac, consultez [Déboguer des compléments Office sur un Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

## <a name="review-the-build-and-debug-properties"></a>Réviser les propriétés de création et de débogage

Avant de commencer le débogage, passez en revue les propriétés de chaque projet pour vérifier que Visual Studio ouvre l’application Office souhaitée et que les autres propriétés de build et de débogage sont définies de manière appropriée.

### <a name="add-in-project-properties"></a>Propriétés du projet de complément

Ouvrez la fenêtre **Propriétés** du projet de complément pour passer en revue les propriétés du projet.

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet de complément.

|Propriété|Description|
|:-----|:-----|
|**Action de démarrage**|Spécifie le mode de débogage pour votre complément. Cette option doit être définie sur **Microsoft Edge** pour un complément Outlook. Pour toutes les autres applications Office, elle doit être définie sur **Office Desktop Client**.|
|**Document de démarrage**<br/>(Compléments Excel, PowerPoint et Word uniquement)|Spécifie le document à ouvrir lors du démarrage du projet. Dans un nouveau projet, il est défini sur **[Nouveau classeur Excel]**, **[Nouveau document Word]** ou **[Nouvelle présentation PowerPoint]**. Pour spécifier un document particulier, suivez les étapes décrites dans [Utiliser un document existant pour déboguer le complément](#use-an-existing-document-to-debug-the-add-in).|
|**Projet Web**|Spécifie le nom du projet web associé au complément.|
|**Adresse e-mail**<br/>(Compléments Outlook uniquement)|Spécifie l’adresse de messagerie du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook. Si vous restez vide, vous êtes invité à entrer l’adresse e-mail lorsque vous démarrez le débogage.|
|**Url EWS**<br/>(Compléments Outlook uniquement)|Spécifie l’URL des services web Exchange (par exemple : `https://www.contoso.com/ews/exchange.aspx`). Cette propriété peut être laissée vide.|
|**Url OWA**<br/>(Compléments Outlook uniquement)|Spécifie l’URL Outlook sur le web (par exemple : `https://www.contoso.com/owa`). Cette propriété peut être laissée vide.|
|**Utiliser l’authentification multi-facteur**<br/>(Compléments Outlook uniquement)|Spécifie la valeur booléenne qui indique si l’authentification multifacteur doit être utilisée. La valeur par défaut est **false**, mais la propriété n’a aucun effet pratique. Si vous devez normalement fournir un deuxième facteur pour vous connecter au compte de messagerie, vous serez invité à le faire lorsque vous commencerez le débogage. |
|**Nom d'utilisateur**<br/>(Compléments Outlook uniquement)|Spécifie le nom du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook. Cette propriété peut être laissée vide.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Précise l’emplacement du fichier de projet.|

> [!NOTE]
> Pour un complément Outlook, vous pouvez choisir de spécifier des valeurs pour une ou plusieurs des propriétés du *complément Outlook uniquement* dans la fenêtre **propriétés**, mais cette opération n’est pas obligatoire.

### <a name="web-application-project-properties"></a>Propriétés du projet application Web

Ouvrez la fenêtre **Propriétés** du projet d’application web pour passer en revue les propriétés du projet.

1. Dans **Explorateur de solutions**, choisissez le projet d’application web.

2. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet d’application web qui sont les plus pertinentes aux projets complément Office.

|Propriété|Description|
|:-----|:-----|
|**SSL activé**|Spécifie si SSL est activé sur le site. Cette propriété doit être définie sur **vrai** pour les projets complément Office.|
|**URL SSL**|Spécifie l’URL HTTPS sécurité pour le site. Lecture seule.|
|**URL**|Spécifie l’URL HTTP pour le site. Lecture seule.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Précise l’emplacement du fichier de projet. Lecture seule. Le fichier manifeste créé par Visual Studio lors de l’exécution est écrit le `bin\Debug\OfficeAppManifests` dossier dans cet emplacement.|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>Déboguer un projet de complément Excel, PowerPoint ou Word

Cette section explique comment démarrer et déboguer un complément Excel, PowerPoint ou Word.

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>Démarrer le projet de complément Excel, PowerPoint ou Word

Démarrez le projet en choisissant **Debug** > **Start Debugging** dans la barre de menus ou appuyez sur le bouton F5. Visual Studio génère automatiquement la solution et démarre l’application hôte Office.

Quand Visual Studio génère le projet, il effectue les tâches suivantes :

1. Crée une copie du fichier manifeste XML et l’ajoute au  `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire. L’application Office qui héberge votre complément consomme cette copie lorsque vous démarrez Visual Studio et déboguez le complément.

2. Crée un ensemble d’entrées de Registre sur votre ordinateur Windows qui permet au complément d’apparaître dans l’application Office.

3. Génère le projet d’application web, puis le déploie sur le serveur web IIS local (`https://localhost`).

4. S’il s’agit du premier projet de complément que vous avez déployé sur le serveur web IIS local, vous pouvez être invité à installer un certificat Self-Signed dans le magasin de certificats racines approuvés de l’utilisateur actuel. Cela est nécessaire pour qu’IIS Express puisse afficher correctement le contenu de votre complément.

> [!NOTE]
> Si Office utilise le contrôle Edge Legacy webview (EdgeHTML) pour exécuter des compléments sur votre ordinateur Windows, Visual Studio peut vous inviter à ajouter une exemption de bouclage de réseau local. Cela est nécessaire pour que le contrôle webview puisse accéder au site web déployé sur le serveur web IIS local. Vous pouvez également modifier ce paramètre à tout moment dans Visual Studio sous **Outils** > **Options** > **Outils Office (web)** > **Débogage de compléments web**. Pour savoir quel contrôle de navigateur est utilisé sur votre ordinateur Windows, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

Visual Studio effectue ensuite les actions suivantes :

1. Modifie l’élément [SourceLocation](/javascript/api/manifest/sourcelocation) du fichier manifeste XML (qui a été copié dans le `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire) en remplaçant le `~remoteAppUrl` jeton par l’adresse complète de la page de démarrage (par exemple, `https://localhost:44302/Home.html`).

2. Il démarre le projet d’application web dans IIS Express.

3. Valide le manifeste. Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md). 

   > [!IMPORTANT]
   > Les fichiers XSD de manifeste Office installés par Visual Studio sont obsolètes. Si vous obtenez des erreurs de validation pour le manifeste, votre première étape de résolution des problèmes doit consister à remplacer un ou plusieurs de ces fichiers par les dernières versions. Pour obtenir des instructions détaillées, consultez [les erreurs de validation de schéma de manifeste dans les projets Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

4. Ouvre l’application Office et charge de manière indépendante votre complément.

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>Déboguer le complément Excel, PowerPoint ou Word

1. Lancez le complément dans l’application Office. Par exemple, s’il s’agit d’un complément du volet Office, il aura ajouté un bouton au ruban **Accueil** (par exemple, un bouton **Afficher le volet Office** ). Sélectionnez le bouton dans le ruban. 

   > [!NOTE]
   > Si votre complément n’est pas chargé de manière indépendante par Visual Studio, vous pouvez le charger manuellement. Dans Excel, PowerPoint ou Word, choisissez l’onglet **Insertion** , puis la flèche vers le bas située à droite de **Mes compléments**.
   >
   > ![Capture d’écran montrant le ruban Insérer dans Excel sur Windows avec la flèche Mes compléments mise en évidence.](../images/excel-cf-register-add-in-1b.png)
   >
   > Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez votre complément pour effectuer cette opération.

   > [!TIP]
   > Le volet Office peut apparaître vide lorsqu’il s’ouvre pour la première fois. Si c’est le cas, il doit s’afficher correctement lorsque vous lancez les outils de débogage dans une étape ultérieure.

3. Ouvrez le [menu personnalité](../design/task-pane-add-ins.md#personality-menu) , puis **choisissez Attacher un débogueur**. Cela ouvre les outils de débogage pour le contrôle webview qu’Office utilise pour exécuter des compléments sur votre ordinateur Windows. Vous pouvez définir des points d’arrêt et parcourir le code comme décrit dans l’un des articles suivants :

    - [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. Pour apporter des modifications à votre code, arrêtez d’abord la session de débogage dans Visual Studio et fermez l’application Office. Apportez vos modifications et démarrez une nouvelle session de débogage.

## <a name="debug-an-outlook-add-in-project"></a>Déboguer un projet de complément Outlook

Cette section explique comment démarrer et déboguer un complément Outlook.

### <a name="start-the-outlook-add-in-project"></a>Démarrer le projet de complément Outlook

Démarrez le projet en choisissant **Debug** > **Start Debugging** dans la barre de menus ou appuyez sur le bouton F5. Visual Studio génère automatiquement la solution et lance la page Outlook de votre location Microsoft 365.

Lorsque Visual Studio génère le projet, il effectue les tâches suivantes.

1. Vous invite à entrer des informations d’identification de connexion. Si vous êtes invité à vous connecter à plusieurs reprises ou si vous recevez une erreur indiquant que vous n’êtes pas autorisé, l’authentification de base peut être désactivée pour les comptes sur votre locataire Microsoft 365. Dans ce cas, essayez d’utiliser un compte Microsoft à la place. Vous pouvez également essayer de définir la propriété **Utiliser l’authentification multifacteur** sur **True** dans le volet des propriétés du projet de complément web Outlook. Consultez les [propriétés du projet de complément](#add-in-project-properties).

1. Crée une copie du fichier manifeste XML et l’ajoute au `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire. Outlook utilise cette copie lorsque vous démarrez Visual Studio et déboguez le complément.

2. Génère le projet d’application web, puis le déploie sur le serveur web IIS local (`https://localhost`).

3. S’il s’agit du premier projet de complément que vous avez déployé sur le serveur web IIS local, vous pouvez être invité à installer un certificat Self-Signed dans le magasin de certificats racines approuvés de l’utilisateur actuel. Cela est nécessaire pour qu’IIS Express puisse afficher correctement le contenu de votre complément.

> [!NOTE]
> Si Office utilise le contrôle Edge Legacy webview (EdgeHTML) pour exécuter des compléments sur votre ordinateur Windows, Visual Studio peut vous inviter à ajouter une exemption de bouclage de réseau local. Cela est nécessaire pour que le contrôle webview puisse accéder au site web déployé sur le serveur web IIS local. Vous pouvez également modifier ce paramètre à tout moment dans Visual Studio sous **Outils** > **Options** > **Outils Office (web)** > **Débogage de compléments web**. Pour savoir quel contrôle de navigateur est utilisé sur votre ordinateur Windows, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

Visual Studio effectue ensuite les actions suivantes :

1. Modifie l’élément [SourceLocation](/javascript/api/manifest/sourcelocation) du fichier manifeste XML (qui a été copié dans le `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire) en remplaçant le `~remoteAppUrl` jeton par l’adresse complète de la page de démarrage (par exemple, `https://localhost:44302/Home.html`).

2. Il démarre le projet d’application web dans IIS Express.

3. Valide le manifeste. Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md). 

   > [!IMPORTANT]
   > Les fichiers XSD de manifeste Office installés par Visual Studio sont obsolètes. Si vous obtenez des erreurs de validation pour le manifeste, votre première étape de résolution des problèmes doit consister à remplacer un ou plusieurs de ces fichiers par les dernières versions. Pour obtenir des instructions détaillées, consultez [les erreurs de validation de schéma de manifeste dans les projets Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

4. Ouvre la page Outlook de votre location Microsoft 365 dans Microsoft Edge.

### <a name="debug-the-outlook-add-in"></a>Déboguer le complément Outlook

1. Dans la page Outlook, sélectionnez un e-mail ou un élément de rendez-vous pour l’ouvrir dans sa propre fenêtre. 

2. Appuyez sur F12 pour ouvrir l’outil de débogage Edge.

3. Une fois l’outil ouvert, lancez le complément. Par exemple, dans la barre d’outils située en haut d’un message, sélectionnez le bouton **Autres applications** , puis sélectionnez votre complément dans la légende qui s’ouvre.

   ![Capture d’écran montrant le bouton Plus d’applications et la légende qu’il ouvre avec le nom et l’icône du complément visibles, ainsi que d’autres icônes d’application.](../images/outlook-more-apps-button.png)

4. Suivez les instructions de l’un des articles suivants pour définir des points d’arrêt et parcourir le code. Ils ont chacun un lien vers des conseils plus détaillés.

   - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > Pour déboguer le code qui s’exécute dans la `Office.initialize` fonction ou une `Office.onReady` fonction qui s’exécute lorsque le complément s’ouvre, définissez vos points d’arrêt, puis fermez et rouvrez le complément. Pour plus d’informations sur ces fonctions, consultez [Initialiser votre complément Office](../develop/initialize-add-in.md).

5. Pour apporter des modifications à votre code, arrêtez d’abord la session de débogage dans Visual Studio et fermez les pages Outlook. Apportez vos modifications et démarrez une nouvelle session de débogage.

## <a name="use-an-existing-document-to-debug-the-add-in"></a>Utiliser un document existant pour déboguer le complément

Si vous avez un document qui contient les données de test à utiliser pendant le débogage de votre complément Excel, PowerPoint ou Word, Visual Studio peut être configuré pour ouvrir ce dernier lorsque vous démarrez le projet. Pour spécifier un document existant à utiliser pour déboguer le complément, procédez comme suit.

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menus, sélectionnez **Project** > **ajouter un élément existant**.

3. Dans la boîte de dialogue **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.

4. Choisissez le bouton **Ajouter** pour ajouter le document à votre projet.

5. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

6. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

7. Dans la fenêtre **propriétés**, choisissez la liste **Document de démarrage** et sélectionnez le document que vous avez ajouté au projet. Le projet est désormais configuré pour démarrer le complément dans ce document.

## <a name="next-steps"></a>Étapes suivantes

Une fois que votre complément fonctionne comme vous le souhaitez, voir [Déployer et publier votre complément Office](../publish/publish.md) pour en savoir plus sur les méthodes avec lesquelles vous pouvez distribuer le complément aux utilisateurs.
