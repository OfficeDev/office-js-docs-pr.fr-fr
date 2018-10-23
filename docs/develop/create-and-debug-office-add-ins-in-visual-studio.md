---
title: Créer et déboguer des compléments Office dans Visual Studio
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 224a4781b894e9bf165d279c30ca16d18bea956d
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681839"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Créer et déboguer des compléments Office dans Visual Studio

Cet article explique comment utiliser Visual Studio pour créer votre premier complément Office. Les étapes décrites dans cet article concernent Visual Studio 2017. Si vous utilisez une autre version de Visual Studio, les procédures peuvent légèrement varier.

> [!NOTE]
> Si vous débutez avec les compléments pour OneNote, reportez-vous à [Créer votre premier complément OneNote](../onenote/onenote-add-ins-getting-started.md).

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Créer un projet de complément Office dans Visual Studio


Pour commencer, vérifiez que les [outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx) sont installés et que vous disposez d’une version de Microsoft Office. Vous pouvez vous joindre au [Programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program), ou suivre ces instructions pour obtenir la [version la plus récente](../develop/install-latest-office-version.md).

1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.
2. Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis sélectionnez un des projets de compléments.
3. Nommez le projet, puis cliquez sur **OK** pour créer le projet.

Dans Visual Studio 2017, les modèles de projet suivants ont des choix supplémentaires une fois que vous cliquez sur **OK** :

**PowerPoint**
- Vous pouvez choisir d’ **Ajouter de nouvelles fonctionnalités dans PowerPoint**, ce qui crée un complément de volet Office.
- Ou vous pouvez choisir d’ **Insérer du contenu dans les diapositives PowerPoint**, ce qui crée un complément de contenu.

**Excel** 
- Vous pouvez choisir d’ **Ajouter de nouvelles fonctionnalités dans Excel**, ce qui crée un complément de volet Office.
- Ou vous pouvez choisir d’ **Insérer du contenu dans la feuille de calcul Excel**, ce qui crée un complément de contenu.
    - Si vous créez un complément de contenu, vous avez un choix supplémentaire : **Complément de base**, qui crée un projet de complément de contenu avec un code de démarrage minimal.
    - Ou vous pouvez choisir un **Complément de visualisation de documents** qui inclut le code de démarrage pour visualiser et lier des données.

Après avoir terminé l’assistant, Visual Studio crée une solution qui contient deux projets. Vous verrez ouvrir la page Home.html par défaut.

|**Projet**|**Description**|
|:-----|:-----|
|Projet de complément|Contient seulement un fichier de manifeste XML, qui contient tous les paramètres qui décrivent votre complément. Ces paramètres aident l’hôte Office à déterminer quand votre complément doit être activé et où il doit apparaître. Visual Studio génère le contenu de ce fichier pour vous afin que vous puissiez exécuter le projet et utiliser immédiatement votre complément. Vous pouvez modifier ces paramètres à tout moment à l’aide de l’éditeur de manifeste.|
|Projet d’application Web|Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur local IIS. Lorsque vous êtes prêt à la publier, vous devez trouver un serveur pour héberger ce projet.Pour en savoir plus sur les projets d’applications web ASP.NET, voir [Projets web ASP.NET](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## <a name="modify-your-add-in-settings"></a>Modifier les paramètres de votre complément


Pour modifier les paramètres de votre complément, modifiez le fichier manifeste XML du projet. Dans l’**Explorateur de solutions**, développez le nœud de projet du complément et le dossier contenant le manifeste XML, puis sélectionnez le manifeste XML. Vous pouvez pointer sur n’importe quel élément du fichier pour afficher une info-bulle qui décrit l’objectif de l’élément. Pour plus d’informations sur le fichier manifeste, voir l’article sur le [manifeste XML de compléments Office](../develop/add-in-manifests.md).


## <a name="develop-the-contents-of-your-add-in"></a>Développer le contenu de votre complément

Alors que le projet de complément vous permet de modifier les paramètres qui décrivent le complément, l’application Web fournit le contenu qui apparaît dans le complément. 

Le projet d’application web contient une page HTML par défaut et le fichier JavaScript que vous pouvez utiliser pour commencer. Ces fichiers contiennent des références aux autres bibliothèques JavaScript, y compris l’API JavaScript pour Office. Vous pouvez développer votre complément par en mettant à jour ces fichiers et en ajoutant plusieurs fichiers HTML et JavaScript. Le tableau suivant décrit les fichiers HTML et JavaScript par défaut.

> [!NOTE]
> Les fichiers dans le tableau ci-dessous peuvent être dans le dossier racine du projet Web, ou dans le dossier **Home** en fonction du type de modèle de projet que vous avez utilisé.

|**Fichier**|**Description**|
|:-----|:-----|
|**Home.html**|La page HTML par défaut du complément. Cette page s’affiche en tant que la première page à l’intérieur du complément lorsqu’il est activé dans un document, un message électronique ou un élément de rendez-vous. Ce fichier contient toutes les références de fichier dont vous avez besoin pour commencer. Vous pouvez commencer à développer votre complément en ajoutant le code HTML à ce fichier.|
|**Home.js**|Le fichier JavaScript associé à la page Home.html. Vous pouvez placer n’importe quel code spécifique au comportement de la page Home.html dans le fichier Home.js. Le fichier Home.js contient des exemples de code pour vous aider.|
|**Home.css**|Définit les styles par défaut à appliquer à votre complément. Nous recommandons l’utilisation de la structure de l’interface utilisateur Office pour la conception et les styles. Pour plus d’informations, voir [Office UI Fabric dans les compléments Office](../design/office-ui-fabric.md).|

> [!NOTE]
> Vous n’êtes pas obligé d’utiliser ces fichiers. N’hésitez pas à ajouter d’autres fichiers au projet et à les utiliser à la place. Si vous souhaitez voir apparaître un autre fichier HTML comme page initiale du complément, ouvrez l’éditeur de manifeste et définissez la propriété **SourceLocation** sur le nom du fichier.

## <a name="debug-your-add-in"></a>Déboguer votre complément

Visual Studio fournit des propriétés de génération et de débogage pour faciliter le débogage de votre complément.

### <a name="review-the-build-and-debug-properties"></a>Revue des propriétés de génération et de débogage

Avant de démarrer la solution, assurez-vous que Visual Studio va ouvrir l’application hôte souhaitée. Cette information apparaît dans les pages de propriétés du projet avec d’autres propriétés liées à la génération et au débogage du complément.

### <a name="to-open-the-property-pages-of-a-project"></a>Ouvrir les pages de propriétés d’un projet

1. Dans l’**Explorateur de solutions**, choisissez le projet de complément de base (et non le projet Web).    
2. Dans la barre de menus, choisissez **Affichage** >  **Fenêtre de propriétés**.
    
Le tableau suivant décrit les propriétés du projet.



|**Propriété**|**Description**|
|:-----|:-----|
|**Action de démarrage**|Indique si votre complément doit être débogué dans un client de bureau Office ou dans un client Office Online dans le navigateur spécifié.|
|**Document de démarrage** (compléments de contenu et du volet Office uniquement)|Spécifie le document à ouvrir lors du démarrage du projet.|
|**Projet Web**|Spécifie le nom du projet Web associé au complément.|
|**Adresse e-mail** (compléments Outlook uniquement)|Spécifie l’adresse e-mail du compte d’utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**URL EWS** (compléments Outlook uniquement)|URL de service Web Exchange (par exemple : https://www.contoso.com/ews/exchange.aspx). |
|**URL OWA** (compléments Outlook uniquement)|URL d’application Web Outlook (par exemple : https://www.contoso.com/owa).|
|**Nom d’utilisateur** (compléments Outlook uniquement)|Spécifie le nom de votre compte d’utilisateur dans Exchange Server ou Exchange Online.|
|**Fichier projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Emplacement du fichier projet.|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>Utiliser un document existant pour déboguer le complément (compléments de contenu et du volet Office uniquement)

Vous pouvez ajouter des documents au projet de complément. Si vous disposez d’un document qui contient des données de test que vous souhaitez utiliser avec votre application, Visual Studio ouvre ce document lorsque vous commencez le projet.

### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>Pour utiliser un document existant pour déboguer le complément

1. Dans l’ **Explorateur de solutions**, choisissez le dossier du projet de complément.
    
    > [!NOTE]
    > Choisissez le projet de complément et non le projet d’application Web.

2. Dans le menu **Projet**, choisissez **Ajouter un élément existant**.
    
3. Dans la boîte de dialogue **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.
    
4. Choisissez le bouton **Ajouter** pour ajouter le document à votre projet.
    
5. Dans l’ **Explorateur de solutions**, choisissez le dossier du projet de complément.
6. Dans la barre de menus, choisissez **Affichage** > **Fenêtre de propriétés**.
7. Dans la fenêtre Propriétés, sélectionnez la liste **Document de démarrage** , puis choisissez le document que vous avez ajouté au projet. Maintenant, le projet est configuré pour démarrer votre complément dans votre document existant.

### <a name="start-the-solution"></a>Démarrer la solution

Démarrez la solution à partir de la barre de menus en choisissant **Déboguer** > **Démarrer le débogage**. Visual Studio génère automatiquement la solution et démarre Office pour héberger votre complément.

Lorsque Visual Studio génère le projet, il effectue les tâches suivantes :

1. Il crée une copie du fichier manifeste XML et l’ajoute au répertoire _NomDuProjet_\Output. L’application hôte utilise cette copie lorsque vous démarrez Visual Studio et déboguez l’application.
    
2. Il crée un ensemble d’entrées dans le Registre de votre ordinateur qui permettent au complément d’apparaître dans l’application hôte.
    
3. Il génère le projet d’application web, puis le déploie sur le serveur web IIS local (http://localhost). 
    
Visual Studio effectue ensuite les actions suivantes :

1. Il modifie l’élément [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js) du fichier manifeste XML en remplaçant le jeton ~remoteAppUrl par l’adresse complète de la page de démarrage (par exemple, http://localhost/MyAgave.html).
    
2. Il démarre le projet d’application Web dans IIS Express.
    
3. Il ouvre l’application hôte. 
    
Visual Studio n’affiche pas les erreurs de validation dans la fenêtre  **OUTPUT** lorsque vous générez le projet. Visual Studio signale au fur et à mesure les erreurs et les avertissements dans la fenêtre **ERRORLIST**. Visual Studio signale également les erreurs de validation en affichant des traits de soulignement ondulés (appelés aussi zigzags) de différentes couleurs dans le code et l’éditeur de texte. Ces marques sont là pour vous indiquer les problèmes détectés par Visual Studio dans votre code. Pour plus d’informations, voir la page relative au [Code et Éditeur de texte](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). Pour plus d’informations sur l’activation ou la désactivation de la validation, voir les rubriques suivantes : 

- [Options, Éditeur de texte, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [Procédure : définir des options de validation pour l’édition HTML dans Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, voir Validation, CSS, Éditeur de texte, Boîte de dialogue Options](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a>Afficher un complément dans Excel ou Word, et parcourir votre code

Si vous définissez la propriété **Document de démarrage** du projet de complément pour Excel ou Word, Visual Studio crée un nouveau document et le complément apparaît. Si vous définissez la propriété **Document de démarrage** du projet de complément pour utiliser un document existant, Visual Studio ouvre le document, mais vous devez insérer le complément manuellement.

1. Dans Excel ou Word, sous l’onglet **Insérer** , choisissez la zone de liste déroulante **Mes compléments** . Sélectionnez la liste à partir de la flèche déroulante, non le bouton lui-même qui ouvre la boîte de dialogue **Compléments Office**.
2. Sous **Compléments pour les développeurs**, choisissez votre complément.

Dans Visual Studio, vous pouvez définir des points d’arrêt et interagir avec votre complément et parcourir le code dans vos fichiers HTML ou JavaScript.

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Afficher le complément Outlook dans Outlook et parcourir votre code

Pour voir le complément dans Outlook, ouvrez un message électronique ou un élément de rendez-vous.

Outlook active le complément pour l’élément à condition que les critères d’activation soient respectés. La barre complément apparaît en haut de la fenêtre de l’inspecteur ou du volet de lecture, et votre complément Outlook apparaît sous la forme d’un bouton dans la barre du complément. Si votre complément est doté d’une commande, un bouton apparaît dans le ruban (soit dans l’onglet par défaut, soit dans un onglet personnalisé indiqué), et le complément n’apparaît pas dans la barre complément.

Pour voir votre complément Outlook, cliquez sur le bouton correspondant.

Dans Visual Studio, vous pouvez définir des points d’arrêt et interagir avec votre complément et parcourir le code dans vos fichiers HTML ou JavaScript.

Vous pouvez également modifier votre code et vérifier les effets de ces modifications dans votre complément Outlook sans devoir fermer le complément Office ni redémarrer le projet. Dans Outlook, ouvrez simplement le menu contextuel du complément Outlook, puis choisissez **Recharger**.


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>Modifier le code et continuer le débogage du complément sans redémarrer le projet

Vous pouvez modifier votre code et passer en revue les effets de ces modifications dans votre complément sans avoir à fermer l’application hôte et redémarrer le projet. Une fois que vous modifiez et enregistrez votre code, ouvrez le menu contextuel pour le complément, puis choisissez **Recharger**.
    

## <a name="next-steps"></a>Étapes suivantes

- [Déploiement et publication de votre complément Office](../publish/publish.md)
    
