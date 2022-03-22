---
title: Office glossaire des termes des compléments
description: Glossaire des termes couramment utilisés dans la documentation Office compléments.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 327c7a8bcc8c3ab21c437c50003e57d34fb933e0
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711216"
---
# <a name="office-add-ins-glossary"></a>Office glossaire des compléments

Il s’agit d’un glossaire des termes couramment utilisés dans la documentation Office compléments.

## <a name="add-in"></a>macro complémentaire

Office sont des applications web qui étendent Office applications. Ces applications web ajoutent de nouvelles fonctionnalités à l’application Office, telles que l’ajout de données externes, l’automatisation des processus ou l’incorporation d’objets interactifs dans Office documents.

Office les applications sont différentes de VBA, COM et VSTO, car elles offrent une prise en charge sur plusieurs plateformes (généralement web, Windows, Mac et iPad) et sont basées sur des technologies web standard (HTML, CSS et JavaScript). Le langage de programmation principal d’un Office est JavaScript ou TypeScript.

## <a name="add-in-commands"></a>commandes de add-in

**Les commandes de add-in** sont des éléments d’interface utilisateur, tels que des boutons et des menus, qui étendent Office’interface utilisateur de votre module. Lorsque les utilisateurs sélectionnent un élément de commande de add-in, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage du module dans un volet De tâches. Les commandes de votre Office, ce qui donne aux utilisateurs davantage de confiance dans votre add-in. Pour en savoir plus, consultez les commandes de Excel[, PowerPoint, word](../design/add-in-commands.md) et de Outlook pour [](../outlook/add-in-commands-for-outlook.md) en savoir plus.

Voir aussi : [ruban, bouton du ruban](#ribbon-ribbon-button).

## <a name="application"></a>application

**L’application** fait référence à une Office application. Les applications Office qui Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [client](#client), [hôte](#host), [Office application, Office client](#office-application-office-client).

## <a name="application-specific-api"></a>API spécifique à l’application

Les API spécifiques à l’application fournissent des objets fortement typés qui interagissent avec des objets natifs d’une application Office spécifique. Par exemple, vous appelez les API JavaScript Excel pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques, etc. Les API spécifiques à l’application sont actuellement disponibles pour Excel, OneNote, PowerPoint, Visio et Word. Pour en savoir [plus, consultez le modèle d’API](../develop/application-specific-api-model.md) propre à l’application.

Voir aussi : [API communes](#common-api).

## <a name="client"></a>client

**Le client** fait généralement référence à une application Office client. Les applications Office, ou clients, qui prisent en charge les Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [hôte](#host), [Office application, Office client](#office-application-office-client).

## <a name="common-api"></a>API communes

Les API communes sont utilisées pour accéder à des fonctionnalités telles que l’interface utilisateur, les boîtes de dialogue et les paramètres clients qui sont communs à plusieurs applications Office client. Ce modèle d’API utilise des[rappels](https://developer.mozilla.org/docs/Glossary/Callback_function), qui vous permettent de spécifier une seule opération dans chaque demande envoyée à l’application Office.

Les API courantes ont été introduites avec Office 2013 et sont utilisées pour interagir avec Office 2013 ou ultérieure. Certaines API communes sont des API héritées du début des années 2010. Excel, PowerPoint et Word ont tous des fonctionnalités d’API communes, mais la plupart de ces fonctionnalités ont été remplacées ou remplacées par le modèle d’API propre à l’application. Les API spécifiques à l’application sont privilégiées lorsque cela est possible.

D’autres API communes, telles que les API communes liées à Outlook, l’interface utilisateur et l’authentification, sont les API modernes et préférées à ces fins. Pour plus d’informations sur le modèle objet d’API commune, voir [Modèle d’objet API JavaScript commun](../develop/office-javascript-api-object-model.md).

Voir aussi : [API spécifique à l’application](#application-specific-api).

## <a name="content-add-in"></a>add-in de contenu

**Les add-ins de** contenu sont des vues web ou des vues de navigateur web qui sont incorporées directement dans Excel, OneNote ou PowerPoint documents. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document. Pour [plus d’informations, voir Office de contenu](../design/content-add-ins.md).

Voir aussi : [webview](#webview).

## <a name="content-delivery-network-cdn"></a>réseau de distribution de contenu (CDN)

Un **réseau ou un CDN** de distribution de contenu est un réseau distribué de serveurs et de centres de données. Il offre généralement une disponibilité et des performances des ressources plus élevées par rapport à un seul serveur ou centre de données.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (également appelée Contoso et Contoso University) est une société fictive utilisée par Microsoft comme exemple de société et de domaine.

## <a name="custom-function"></a>fonction personnalisée

Une **fonction personnalisée est** une fonction définie par l’utilisateur qui est empaqueté avec un Excel de commande. Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions, au-delà des fonctionnalités Excel classiques, en définissant ces fonctions dans JavaScript dans le cadre d’un add-in. Les utilisateurs Excel peuvent accéder aux fonctions personnalisées de la même façon que n’importe quelle fonction native dans Excel. Voir [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md) pour en savoir plus.

## <a name="custom-functions-runtime"></a>runtime de fonctions personnalisées

Un **runtime de fonctions personnalisées est un runtime** JavaScript qui exécute uniquement des fonctions personnalisées. Il n’a pas d’interface utilisateur et ne peut pas interagir Office.js API. Si votre add-in ne possède que des fonctions personnalisées, il s’agit d’un runtime léger à utiliser. Si vos fonctions personnalisées doivent interagir avec le volet Des tâches ou Office.js API, configurez un runtime JavaScript partagé. Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Voir aussi : [runtime JavaScript](#javascript-runtime), [runtime JavaScript partagé, runtime partagé](#shared-javascript-runtime-shared-runtime).

## <a name="host"></a>host

**L’hôte** fait généralement référence à une application Office’application. Les applications Office, ou hôtes, qui prisent en charge les Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [Office application, Office client](#office-application-office-client).

## <a name="javascript-runtime"></a>Runtime JavaScript

Le **runtime JavaScript est** l’environnement hôte du navigateur dans qui s’exécute le add-in. Dans Office sur Windows et Office sur Mac, le runtime JavaScript est un contrôle de navigateur incorporé (ou webview) tel qu’Internet Explorer, Edge Legacy, Edge WebView2 ou Safari. Différentes parties d’un add-in s’exécutent dans des runtimes JavaScript distincts. Par exemple, les commandes de add-in, les fonctions personnalisées et le code du volet Des tâches utilisent généralement des runtimes JavaScript distincts, sauf si vous configurez un runtime JavaScript partagé. Pour [plus d’informations, consultez les navigateurs utilisés Office les modules complémentaires](../concepts/browsers-used-by-office-web-add-ins.md).

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime JavaScript partagé, runtime partagé](#shared-javascript-runtime-shared-runtime), [webview](#webview).

## <a name="office-application-office-client"></a>Office application, client Office client

**Office client fait** référence à une application Office client. Les applications Office, ou clients, qui prisent en charge les Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [hôte](#host).

## <a name="platform"></a>platform

Une **plateforme** fait généralement référence au système d’exploitation qui exécute l Office application. Les plateformes qui Office les Windows, Mac, iPad et les navigateurs web.

## <a name="quick-start"></a>démarrage rapide

Un **démarrage rapide** est une description de haut niveau des compétences clés et des connaissances requises pour le fonctionnement de base d’un programme particulier. Dans la documentation Office de l’application, un démarrage rapide est une introduction au développement d’un module pour une application particulière, telle que Outlook. Un démarrage rapide contient une série d’étapes qu’un développeur de compl?ment peut effectuer en environ 5 minutes, ce qui a pour effet de d finir un environnement de d veloppeur fonctionnel et de compl?ment fonctionnel.

Voir aussi : [didacticiel](#tutorial).

## <a name="requirement-set"></a>ensemble de conditions requises

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>ruban, bouton de ruban

Un **ruban** est une barre de commandes qui organise les fonctionnalités d’une application en une série d’onglets ou de boutons en haut d’une fenêtre. Un **bouton de ruban** est l’un des boutons de cette série. Pour [plus d’informations, voir Afficher ou masquer Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) ruban.

## <a name="runtime"></a>runtime

Voir : [Runtime JavaScript](#javascript-runtime).

## <a name="shared-javascript-runtime-shared-runtime"></a>runtime JavaScript partagé, runtime partagé

Un **runtime JavaScript** partagé, ou **runtime** partagé, permet à tout le code de votre module, y compris le volet Des tâches, les commandes de add-in et les fonctions personnalisées, de s’exécuter dans le même runtime JavaScript et de continuer à s’exécuter même lorsque le volet Des tâches est fermé. Pour en savoir plus, voir Configurer votre Office de Office pour utiliser un [runtime JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md) et un Astuces partagés pour utiliser le [runtime JavaScript](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) partagé dans votre Office.

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime JavaScript](#javascript-runtime).

## <a name="task-pane"></a>volet Office

Les volets Des tâches sont des surfaces d’interface ou des vues web qui apparaissent généralement sur le côté droit de la fenêtre dans Excel, Outlook, PowerPoint et Word. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets Des tâches lorsque vous n’avez pas besoin ou ne pouvez pas incorporer de fonctionnalités directement dans le document. Pour [en savoir plus, consultez les volets Office des tâches](../design/task-pane-add-ins.md).

Voir aussi : [webview](#webview).

## <a name="tutorial"></a>didacticiel

Un **didacticiel** est un aide pédagogique conçu pour aider les personnes à apprendre à utiliser un produit ou une procédure. Dans le Office des applications, un didacticiel guide un développeur de compl?ments tout au long du processus de d veloppe de compl?ment pour une application en particulier, par exemple Excel. Cela implique le suivi de 20 étapes ou plus et est un plus grand investissement en temps [qu’un démarrage rapide](#quick-start).

Voir aussi : [démarrage rapide](#quick-start).

## <a name="ui-less-custom-function"></a>Fonction personnalisée sans interface utilisateur

**Les fonctions personnalisées sans interface utilisateur s’exécutent** dans le runtime des fonctions personnalisées. Ils n’ont pas d’interface utilisateur et ne peuvent pas interagir Office.js API.

Voir aussi : [fonction personnalisée](#custom-function), [runtime de fonctions personnalisées](#custom-functions-runtime).

## <a name="web-add-in"></a>web add-in

**Le add-in web** est un terme hérité pour un Office de recherche. Ce terme peut être utilisé lorsque la documentation Microsoft 365 doit faire la distinction entre les Office modernes et d’autres types de compl?ments tels que VBA, COM ou VSTO.

Voir aussi : [add-in](#add-in).

## <a name="webview"></a>webview

Une **vue web est** un élément ou une vue qui affiche du contenu web à l’intérieur d’une application. Les add-ins de contenu et les volets De tâches contiennent tous deux des navigateurs web incorporés et sont des exemples de vues web dans Office des applications.

Voir aussi : [add-in de contenu](#content-add-in), [volet Des tâches](#task-pane).

## <a name="xll"></a>XLL

Un **Excel XLL** est un fichier de Excel qui fournit des fonctions définies par l’utilisateur et possède l’extension **de fichier .xll**. Un fichier XLL est un type de fichier de bibliothèque de liens dynamiques (DLL) qui ne peut être ouvert qu’Excel. Les fichiers de add-in XLL doivent être écrits en C ou C++. Les fonctions personnalisées sont l’équivalent moderne des fonctions XLL définies par l’utilisateur. Les fonctions personnalisées offrent une prise en charge sur toutes les plateformes et sont à compatibilité avec les fichiers XLL vers l’arrière. Pour [plus d’informations, voir Étendre les fonctions personnalisées avec des fonctions XLL définies par l’utilisateur](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) .

Voir aussi : [fonction personnalisée](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Générateur Yeoman, yo office

Le [générateur Yeoman pour les Office utilise](../develop/yeoman-generator-overview.md) l’outil [open source Yeoman](https://github.com/yeoman/yo) pour générer un Office par le biais de la ligne de commande. `yo office`est la commande qui exécute le générateur Yeoman pour Office de recherche. Les Office des modules et didacticiels utilisent le générateur Yeoman.

## <a name="see-also"></a>Voir aussi

- [Ressources supplémentaires sur les compléments Office](resources-links-help.md)